/**
 * PROMPT-041 — the reduced-fidelity notice must say WHY, and must not lie about the cause.
 *
 * The bug being pinned: while the user's own PowerPoint is open, Report Extractor (correctly, for their data's
 * sake) refuses to attach to it and falls back to the built-in renderer.  The preview then said only
 * "Bản xem trước đơn giản" — no cause and no remedy — so the reviewer had no way to know that closing PowerPoint
 * would restore a high-fidelity preview, and the message was indistinguishable from "this PC has no PowerPoint".
 *
 * Python now sends a stable slug in `slidePreviewReason`.  These tests render the REAL `LearningTab` and assert
 *   • the slug → text mapping keeps those two opposite cases apart;
 *   • the notice stays a single non-blocking banner with the original reduced-fidelity wording intact, and the
 *     actionable sentence is appended (so the PROMPT-027 notice contract is not regressed);
 *   • a genuine PowerPoint render shows no notice at all, hint or otherwise.
 */
import assert from 'node:assert/strict'
import { test } from 'node:test'
import { JSDOM } from 'jsdom'
import { createServer } from 'vite'
import * as React from 'react'
import { fileURLToPath } from 'node:url'

const frontendRoot = fileURLToPath(new URL('..', import.meta.url))
const delay = (ms) => new Promise((resolve) => globalThis.setTimeout(resolve, ms))

const APP_VERSION = '1.3.7'
const APP_BUILD = '020'
const SLIDE_W = 9144000
const SLIDE_H = 5143500
const PREVIEW = 'data:image/svg+xml;base64,PHN2ZyBpZD0ic2xpZGUiLz4='

function emptyLearningState(imageCount = 0) {
  return {
    images: [],
    contents: [],
    candidateStatus: {
      image: { state: 'ready', reason: imageCount ? 'available' : 'empty', count: imageCount,
        message: imageCount ? `Đã tải ${imageCount} ảnh ứng viên từ lượt xử lý hiện tại.` : 'Lượt xử lý hiện tại không có ảnh ứng viên cần kiểm tra.' },
      content: { state: 'ready', reason: 'empty', count: 0, message: 'Lượt xử lý hiện tại không có khối nội dung cần kiểm tra.' },
    },
    counts: { image: { total: 0, labeled: 0 }, content: { total: 0, labeled: 0 } },
    modelStatus: { image: 'Ready', content: 'Ready' },
    excelPending: 0,
    excelLastResult: { kind: 'none', message: '' },
    training: false,
  }
}

function candidate({ backend, faithful, reason }) {
  return {
    id: 'image-1',
    sourceFile: 'improvement-report.pptx',
    managementNumber: 'TNP-001',
    slide: 5,
    pictureId: '42',
    src: 'data:image/png;base64,AA==',
    bounds: { x: 20, y: 25, w: 30, h: 35 },
    slideWidth: SLIDE_W,
    slideHeight: SLIDE_H,
    targetBbox: { x: 4572000, y: 1800810, width: 1097280, height: 720090 },
    targetBboxPct: { x: 20, y: 25, w: 30, h: 35 },
    targetKind: 'picture',
    slidePreview: PREVIEW,
    slidePreviewWidth: 1600,
    slidePreviewHeight: 900,
    slidePreviewBackend: backend,
    slidePreviewFaithful: faithful,
    slidePreviewReason: reason,
    evidenceRegionBbox: { x: 4297680, y: 1597660, width: 3020688, height: 1234440 },
    evidenceRegionBboxPct: { x: 18, y: 22, w: 40, h: 45 },
    evidenceRegionKind: 'after_region',
    evidenceRegionItemId: 'item-one',
    evidenceRegionItemIndex: 0,
    evidenceRegionItemHeading: 'Mục một',
    evidenceRegionPictureCount: 2,
    evidenceRegionPictureIds: [7, 8],
    evidenceRegionId: 'region-1',
    evidenceRegionBlockIndex: 0,
    evidenceRegionBoundarySource: 'heading',
    evidenceRegionNextHeadingTop: null,
    itemId: 'item-one',
    itemIndex: 0,
    itemHeading: 'Mục một',
    decision: 'AFTER',
    confidenceBand: 'high',
    confidence: 0.94,
    evidence: ['Candidate evidence'],
    excelEligible: true,
    userLabel: 'UNLABELED',
    note: '',
  }
}

const envelope = (data) => ({ ok: true, data })

async function mountLearningHarness() {
  const dom = new JSDOM('<!doctype html><html><body><div id="root"></div></body></html>', { url: 'http://localhost/' })
  const globalNames = [
    'window', 'document', 'navigator', 'HTMLElement', 'Node', 'Event', 'MouseEvent',
    'ResizeObserver', 'IS_REACT_ACT_ENVIRONMENT',
  ]
  const priorGlobals = new Map(globalNames.map((name) => [name, Object.getOwnPropertyDescriptor(globalThis, name)]))
  const defineGlobal = (name, value) => Object.defineProperty(globalThis, name, {
    configurable: true, writable: true, value,
  })
  defineGlobal('window', dom.window)
  defineGlobal('document', dom.window.document)
  defineGlobal('navigator', dom.window.navigator)
  defineGlobal('HTMLElement', dom.window.HTMLElement)
  defineGlobal('Node', dom.window.Node)
  defineGlobal('Event', dom.window.Event)
  defineGlobal('MouseEvent', dom.window.MouseEvent)
  defineGlobal('IS_REACT_ACT_ENVIRONMENT', true)

  class TestResizeObserver {
    observe() {}
    unobserve() {}
    disconnect() {}
  }
  defineGlobal('ResizeObserver', TestResizeObserver)
  dom.window.ResizeObserver = TestResizeObserver
  dom.window.HTMLElement.prototype.scrollIntoView = () => {}
  dom.window.matchMedia = (media) => ({
    matches: media === '(min-width: 1280px)', media, onchange: null,
    addEventListener() {}, removeEventListener() {}, addListener() {}, removeListener() {},
    dispatchEvent() { return true },
  })

  const config = {
    paths: { reportFolder: '', template: '', output: '' },
    period: { mode: 'auto', month: '', year: '', from: '', to: '' },
    forceReprocess: false,
    ollama: { host: '', port: '11434', model: '', checked: 'unchecked', message: '', models: [] },
    update: { path: '', autoCheck: true },
  }
  const dashboard = {
    app: { name: 'Report Extractor', title: `Report Extractor v${APP_VERSION}`,
      version: APP_VERSION, build: APP_BUILD, buildNumber: Number(APP_BUILD) },
    config,
    scan: { scanned: false, message: 'No reports selected.' },
    reports: [],
    job: { status: 'idle', queue: [], index: 0, doneCount: 0, stage: 'waiting', percent: 0,
      currentFile: '', elapsedSec: 0, remainSec: null, startedAt: null, finishedAt: null, hasSamples: false },
    logs: [],
    ollama: { host: '', port: '', model: '', checked: 'unchecked', message: '', models: [], aiStatus: '',
      discovering: false, discoveryMessage: '', discoveryResults: [], discoveryChecked: 0,
      discoveryTotal: 0, serverApplyRunning: false, serverApplyMessage: '' },
    update: { path: '', autoCheck: true, status: 'idle', message: '', progress: 0, progressStage: '' },
    diagnostics: { running: false, rows: [], error: '' },
  }
  const state = { learning: emptyLearningState(), version: dashboard.app }

  let server
  let root
  try {
    server = await createServer({
      configFile: `${frontendRoot}/vite.config.ts`, root: frontendRoot,
      appType: 'custom', server: { middlewareMode: true, hmr: false },
    })
    const [{ StoreProvider, useStore }, { LearningTab }] = await Promise.all([
      server.ssrLoadModule('/src/state/store.tsx'),
      server.ssrLoadModule('/src/tabs/LearningTab.tsx'),
    ])
    const { createRoot } = await import('react-dom/client')

    function Harness() {
      const store = useStore()
      return React.createElement(
        React.Fragment,
        null,
        React.createElement('output', { 'data-testid': 'bridge-state' }, store.connected ? 'ready' : 'disconnected'),
        React.createElement('button', {
          type: 'button', 'data-testid': 'refresh-learning',
          onClick: () => { void store.refreshLearning() },
        }, 'Refresh candidates'),
        React.createElement(LearningTab),
      )
    }

    dom.window.pywebview = {
      api: {
        ping: async () => envelope('Python bridge OK'),
        get_app_version: async () => envelope(state.version),
        get_current_config: async () => envelope(config),
        get_dashboard_state: async () => envelope(dashboard),
        get_learning_state: async () => envelope(state.learning),
      },
    }

    root = createRoot(dom.window.document.getElementById('root'))
    await React.act(async () => {
      root.render(React.createElement(StoreProvider, null, React.createElement(Harness)))
    })
    const container = dom.window.document.getElementById('root')
    await React.act(async () => {
      dom.window.dispatchEvent(new dom.window.Event('pywebviewready'))
    })

    const waitFor = async (predicate, description, timeoutMs = 4_000) => {
      const deadline = Date.now() + timeoutMs
      while (Date.now() < deadline) {
        if (predicate()) return
        await React.act(async () => { await delay(10) })
      }
      assert.fail(`Timed out waiting for ${description}`)
    }
    await waitFor(
      () => container.querySelector('[data-testid="bridge-state"]')?.textContent === 'ready',
      'the pywebview startup handshake',
    )

    const click = async (element) => {
      assert.ok(element, 'element to click should exist')
      await React.act(async () => {
        element.dispatchEvent(new dom.window.MouseEvent('click', { bubbles: true }))
        await delay(0)
      })
    }
    const openImageReview = async () => {
      const menu = [...container.querySelectorAll('aside nav button')]
        .find((button) => button.textContent.trim() === 'Kiểm tra ảnh cải tiến')
      await click(menu)
    }
    const setLearning = async (images) => {
      state.learning = {
        ...emptyLearningState(images.length),
        images,
        counts: { image: { total: images.length, labeled: 0 }, content: { total: 0, labeled: 0 } },
      }
      await click(container.querySelector('[data-testid="refresh-learning"]'))
    }
    const panel = () => container.querySelector('section[aria-label="Slide ngữ cảnh"]')
    const notices = () => [...container.querySelectorAll('[role="status"]')]
      .map((node) => node.textContent.replace(/\s+/g, ' ').trim())

    const cleanup = async () => {
      if (root) await React.act(async () => { root.unmount() })
      root = null
      if (server) await server.close()
      server = null
      dom.window.close()
      for (const [name, descriptor] of priorGlobals) {
        if (descriptor) Object.defineProperty(globalThis, name, descriptor)
        else delete globalThis[name]
      }
    }

    return { container, panel, notices, click, openImageReview, setLearning, waitFor, cleanup,
      ssrLoadModule: (id) => server.ssrLoadModule(id) }
  } catch (error) {
    if (root) await React.act(async () => { root.unmount() })
    if (server) await server.close()
    dom.window.close()
    for (const [name, descriptor] of priorGlobals) {
      if (descriptor) Object.defineProperty(globalThis, name, descriptor)
      else delete globalThis[name]
    }
    throw error
  }
}

/* --------------------------------------------------------------------------- the slug → text mapping */
test('the fidelity hint separates "your PowerPoint is open" from "no PowerPoint here"', async () => {
  const harness = await mountLearningHarness()
  try {
    const { fidelityHint } = await harness.ssrLoadModule('/src/tabs/LearningTab.tsx')
    assert.equal(typeof fidelityHint, 'function', 'LearningTab must export fidelityHint')

    const userSession = fidelityHint('powerpoint-user-session-running')
    assert.match(userSession, /Đóng PowerPoint/, 'the user-session case must tell the reviewer what to do')
    assert.match(userSession, /CÓ PowerPoint/i, 'it must NOT imply PowerPoint is missing')
    assert.doesNotMatch(userSession, /không có PowerPoint/)

    const unavailable = fidelityHint('powerpoint-unavailable')
    assert.match(unavailable, /không có PowerPoint/, 'the unavailable case must say PowerPoint is missing')
    assert.doesNotMatch(unavailable, /Đóng PowerPoint/, '...and must not tell them to close a session')
    assert.notEqual(userSession, unavailable, 'the two opposite causes must not collapse into one message')

    // every other refusal slug has its own wording, and a faithful render has none
    for (const code of ['powerpoint-registered-instance', 'powerpoint-process-table-unreadable',
      'powerpoint-ownership-unproven', 'powerpoint-not-exclusive', 'powerpoint-import']) {
      assert.ok(fidelityHint(code).length > 10, `${code} must produce an explanation`)
    }
    assert.equal(fidelityHint(''), '', 'no reason → no extra sentence')
    assert.equal(fidelityHint('some-future-code'), '', 'an unknown slug degrades to the generic notice')
  } finally {
    await harness.cleanup()
  }
})

/* ------------------------------------------------------------------------- the notice in a real render */
test('a blocked-by-user-session preview says so, and stays one non-blocking banner', async () => {
  const harness = await mountLearningHarness()
  try {
    await harness.openImageReview()
    await harness.setLearning([candidate({
      backend: 'builtin', faithful: false, reason: 'powerpoint-user-session-running',
    })])
    await harness.waitFor(() => harness.panel().textContent.includes('Render: Trình vẽ tích hợp'),
      'the builtin backend chip')

    const banners = harness.notices().filter((text) => text.includes('Bản xem trước đơn giản'))
    assert.equal(banners.length, 1, 'exactly one reduced-fidelity banner')
    // the PROMPT-027 wording is intact, the PROMPT-041 cause/remedy is appended to the SAME banner
    assert.match(banners[0], /Bản xem trước đơn giản — bố cục có thể khác PowerPoint/)
    assert.match(banners[0], /Kết quả Excel vẫn dùng đúng toạ độ PPTX\./)
    assert.match(banners[0], /Đóng PowerPoint rồi chọn lại ảnh/)
    // non-blocking: the slide bitmap and the Excel region overlay still render
    assert.ok(harness.panel().querySelector('[data-testid="slide-preview-bitmap"]'), 'the bitmap still renders')
    assert.ok(harness.container.querySelector('[data-testid="overlay-evidence-region"]'), 'the region still renders')
  } finally {
    await harness.cleanup()
  }
})

test('a PowerPoint-missing preview reports the missing install, and a faithful preview warns about nothing', async () => {
  const harness = await mountLearningHarness()
  try {
    await harness.openImageReview()

    await harness.setLearning([candidate({ backend: 'builtin', faithful: false, reason: 'powerpoint-unavailable' })])
    await harness.waitFor(() => harness.panel().textContent.includes('Render: Trình vẽ tích hợp'),
      'the builtin backend chip')
    const missing = harness.notices().filter((text) => text.includes('Bản xem trước đơn giản'))
    assert.equal(missing.length, 1)
    assert.match(missing[0], /không có PowerPoint/)
    assert.doesNotMatch(missing[0], /Đóng PowerPoint/, 'must not tell the user to close a session that is not there')

    await harness.setLearning([candidate({ backend: 'powerpoint', faithful: true, reason: '' })])
    await harness.waitFor(() => harness.panel().textContent.includes('Render: PowerPoint'),
      'the PowerPoint backend chip')
    assert.equal(harness.notices().filter((text) => text.includes('Bản xem trước đơn giản')).length, 0,
      'a genuine PowerPoint render shows no fidelity warning at all')
  } finally {
    await harness.cleanup()
  }
})
