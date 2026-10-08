/**
 * PROMPT-027 — Learning preview overlay truth, honest backend reporting and visible build identity.
 *
 * Covers, in a real React render of the actual `LearningTab` (§13–§19, §40, §41, §44, §50):
 *   • the FINAL Excel evidence region and the picture candidate are rendered as TWO distinct overlays,
 *     each positioned from the authoritative percentages Python sent — React never recomputes them;
 *   • a reduced-fidelity preview backend (builtin/LibreOffice) shows a small NON-BLOCKING notice, while a
 *     genuine PowerPoint render shows none;
 *   • the compact overlay toggle switches geometry without unmounting the component (hook order intact);
 *   • a candidate owned by no Excel region never borrows another item's region;
 *   • the header exposes version AND build so a Windows tester can spot a stale checkout immediately.
 */
import assert from 'node:assert/strict'
import { test } from 'node:test'
import { JSDOM } from 'jsdom'
import { createServer } from 'vite'
import * as React from 'react'
import { fileURLToPath } from 'node:url'

const frontendRoot = fileURLToPath(new URL('..', import.meta.url))
const delay = (ms) => new Promise((resolve) => globalThis.setTimeout(resolve, ms))

const APP_VERSION = '1.3.3'
const APP_BUILD = '016'

const SLIDE_W = 9144000
const SLIDE_H = 5143500
const PREVIEW = 'data:image/svg+xml;base64,PHN2ZyBpZD0ic2xpZGUiLz4='

function emptyLearningState() {
  return {
    images: [],
    contents: [],
    counts: { image: { total: 0, labeled: 0 }, content: { total: 0, labeled: 0 } },
    modelStatus: { image: 'Ready', content: 'Ready' },
    excelPending: 0,
    excelLastResult: { kind: 'none', message: '' },
    training: false,
  }
}

/**
 * One picture candidate plus the authoritative evidence region that owns it.
 * The region is deliberately LARGER than the picture (one item = several After pictures + annotation),
 * which is exactly the distinction §13 requires the UI to make visible.
 */
function candidate({ id, slide, pictureId, itemIndex, targetPct, evidencePct, backend, faithful, evidenceMeta }) {
  return {
    id,
    sourceFile: 'improvement-report.pptx',
    managementNumber: 'TNP-001',
    slide,
    pictureId,
    src: 'data:image/png;base64,AA==',
    bounds: { x: targetPct.x, y: targetPct.y, w: targetPct.w, h: targetPct.h },
    slideWidth: SLIDE_W,
    slideHeight: SLIDE_H,
    targetBbox: { x: 4572000, y: 1800810, width: 1097280, height: 720090 },
    targetBboxPct: targetPct,
    targetKind: 'picture',
    slidePreview: PREVIEW,
    slidePreviewWidth: 1600,
    slidePreviewHeight: 900,
    slidePreviewBackend: backend,
    slidePreviewFaithful: faithful,
    evidenceRegionBbox: evidencePct
      ? { x: 4297680, y: 1597660, width: 3020688, height: 1234440 }
      : null,
    evidenceRegionBboxPct: evidencePct ?? null,
    evidenceRegionKind: evidencePct ? 'after_region' : '',
    evidenceRegionItemId: evidenceMeta?.itemId ?? '',
    evidenceRegionItemIndex: evidenceMeta?.itemIndex ?? -1,
    evidenceRegionItemHeading: evidenceMeta?.itemHeading ?? '',
    evidenceRegionPictureCount: evidenceMeta?.pictureCount ?? 0,
    itemId: evidenceMeta?.itemId ?? `item-${slide}`,
    itemIndex,
    itemHeading: `Improvement item ${slide}`,
    decision: 'AFTER',
    confidenceBand: 'high',
    confidence: 0.94,
    evidence: ['Candidate evidence'],
    excelEligible: true,
    userLabel: 'UNLABELED',
    note: '',
  }
}

async function waitFor(predicate, description, timeoutMs = 4_000) {
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    if (predicate()) return
    await React.act(async () => { await delay(10) })
  }
  assert.fail(`Timed out waiting for ${description}`)
}

const envelope = (data) => ({ ok: true, data })

async function mountLearningHarness() {
  const dom = new JSDOM('<!doctype html><html><body><div id="root"></div></body></html>', {
    url: 'http://localhost/',
  })
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
    matches: media === '(min-width: 1280px)',
    media,
    onchange: null,
    addEventListener() {},
    removeEventListener() {},
    addListener() {},
    removeListener() {},
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
    app: {
      name: 'Report Extractor', title: `Report Extractor v${APP_VERSION}`,
      version: APP_VERSION, build: APP_BUILD, buildNumber: Number(APP_BUILD),
    },
    config,
    scan: { scanned: false, message: 'No reports selected.' },
    reports: [],
    job: {
      status: 'idle', queue: [], index: 0, doneCount: 0, stage: 'waiting', percent: 0,
      currentFile: '', elapsedSec: 0, remainSec: null, startedAt: null, finishedAt: null, hasSamples: false,
    },
    logs: [],
    ollama: {
      host: '', port: '', model: '', checked: 'unchecked', message: '', models: [], aiStatus: '',
      discovering: false, discoveryMessage: '', discoveryResults: [], discoveryChecked: 0,
      discoveryTotal: 0, serverApplyRunning: false, serverApplyMessage: '',
    },
    update: { path: '', autoCheck: true, status: 'idle', message: '', progress: 0, progressStage: '' },
    diagnostics: { running: false, rows: [], error: '' },
  }

  const state = { learning: emptyLearningState(), version: dashboard.app }

  let server
  let root
  try {
    server = await createServer({
      configFile: `${frontendRoot}/vite.config.ts`,
      root: frontendRoot,
      appType: 'custom',
      server: { middlewareMode: true, hmr: false },
    })
    const [{ StoreProvider, useStore }, { LearningTab }, { AppHeader }] = await Promise.all([
      server.ssrLoadModule('/src/state/store.tsx'),
      server.ssrLoadModule('/src/tabs/LearningTab.tsx'),
      server.ssrLoadModule('/src/components/AppHeader.tsx'),
    ])
    const { createRoot } = await import('react-dom/client')

    function Harness() {
      const store = useStore()
      return React.createElement(
        React.Fragment,
        null,
        React.createElement(AppHeader),
        React.createElement('output', { 'data-testid': 'bridge-state' }, store.connected ? 'ready' : 'disconnected'),
        React.createElement('button', {
          type: 'button',
          'data-testid': 'refresh-learning',
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
      root.render(React.createElement(React.StrictMode, null,
        React.createElement(StoreProvider, null, React.createElement(Harness))))
    })
    const container = dom.window.document.getElementById('root')
    await React.act(async () => {
      dom.window.dispatchEvent(new dom.window.Event('pywebviewready'))
    })
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
        ...emptyLearningState(),
        images,
        counts: { image: { total: images.length, labeled: 0 }, content: { total: 0, labeled: 0 } },
      }
      await click(container.querySelector('[data-testid="refresh-learning"]'))
    }
    const panel = () => container.querySelector('section[aria-label="Slide ngữ cảnh"]')
    const overlay = (testId) => container.querySelector(`[data-testid="${testId}"]`)
    const boxOf = (element) => {
      assert.ok(element, 'overlay element should exist')
      const style = element.getAttribute('style') || ''
      const read = (name) => {
        const match = style.match(new RegExp(`${name}:\\s*([0-9.]+)%`))
        return match ? Number(match[1]) : null
      }
      return { left: read('left'), top: read('top'), width: read('width'), height: read('height') }
    }
    const byText = (selector, text) => [...container.querySelectorAll(selector)]
      .find((element) => element.textContent.trim() === text)

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

    return { container, panel, overlay, boxOf, byText, click, openImageReview, setLearning, waitFor, cleanup }
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

/* ------------------------------------------------------------------ §13/§16/§19 two distinct overlays */
test('evidence region and picture candidate render as two distinct, DTO-driven overlays', async () => {
  const harness = await mountLearningHarness()
  try {
    const targetPct = { x: 50, y: 35, w: 12, h: 14 }
    const evidencePct = { x: 47, y: 31, w: 33, h: 24 }        // strictly larger, contains the picture
    await harness.openImageReview()
    await harness.setLearning([candidate({
      id: 'image-1', slide: 5, pictureId: '42', itemIndex: 1, targetPct, evidencePct,
      backend: 'powerpoint', faithful: true,
      evidenceMeta: { itemId: 'item-two', itemIndex: 1, itemHeading: 'Mục hai', pictureCount: 3 },
    })])
    await harness.waitFor(() => harness.overlay('overlay-evidence-region'), 'both overlays')

    const evidence = harness.overlay('overlay-evidence-region')
    const picture = harness.overlay('overlay-picture-candidate')
    assert.ok(evidence && picture, 'both overlay geometries must render')

    // §19/§40: each overlay is positioned with EXACTLY the percentages Python sent — never recomputed
    assert.deepEqual(harness.boxOf(evidence), {
      left: evidencePct.x, top: evidencePct.y, width: evidencePct.w, height: evidencePct.h,
    })
    assert.deepEqual(harness.boxOf(picture), {
      left: targetPct.x, top: targetPct.y, width: targetPct.w, height: targetPct.h,
    })
    assert.notDeepEqual(harness.boxOf(evidence), harness.boxOf(picture))

    // §16: the evidence region is visually dominant (heavier solid border), the candidate is dashed
    assert.match(evidence.className, /border-2/)
    assert.match(evidence.className, /border-solid|border-brand-600/)
    assert.doesNotMatch(evidence.className, /dashed/)
    assert.match(picture.className, /border-dashed/)

    // §16: labels distinguish the two geometries
    assert.match(evidence.textContent, /Mục #2 · Vùng xuất Excel/)
    assert.match(picture.textContent, /Ảnh #42/)
    assert.doesNotMatch(picture.textContent, /Vùng xuất Excel/)

    // §56-F/§41: a genuine PowerPoint render reports its backend and shows NO fidelity warning
    assert.match(harness.panel().textContent, /Render: PowerPoint/)
    assert.doesNotMatch(harness.panel().textContent, /Bản xem trước đơn giản/)

    // the right panel exposes both geometries as plain numbers for Windows acceptance (§42)
    const info = harness.container.querySelector('section[aria-label="Thông tin và nhãn xác nhận"]')
    assert.match(info.textContent, /Vùng xuất Excel/)
    assert.match(info.textContent, /3 ảnh/)
    assert.match(info.textContent, /trung thực PowerPoint/)
    assert.match(info.textContent, /Ảnh #42/)
  } finally {
    await harness.cleanup()
  }
})

/* ------------------------------------------------------------------ §6/§41 honest fallback reporting */
test('a reduced-fidelity backend shows a non-blocking notice and PowerPoint does not', async () => {
  const harness = await mountLearningHarness()
  try {
    await harness.openImageReview()
    const base = {
      id: 'image-1', slide: 5, pictureId: '42', itemIndex: 0,
      targetPct: { x: 20, y: 25, w: 30, h: 35 },
      evidencePct: { x: 18, y: 22, w: 40, h: 45 },
      evidenceMeta: { itemId: 'item-one', itemIndex: 0, itemHeading: 'Mục một', pictureCount: 2 },
    }

    for (const [backend, faithful, expectNotice] of [['builtin', false, true], ['libreoffice', false, true],
      ['powerpoint', true, false]]) {
      await harness.setLearning([candidate({ ...base, backend, faithful })])
      await harness.waitFor(
        () => harness.panel().textContent.includes(`Render: ${backend === 'builtin' ? 'Trình vẽ tích hợp' : backend === 'libreoffice' ? 'LibreOffice' : 'PowerPoint'}`),
        `the ${backend} backend chip`,
      )
      const notice = /Bản xem trước đơn giản — bố cục có thể khác PowerPoint/
      if (expectNotice) {
        assert.match(harness.panel().textContent, notice, `${backend} must warn about reduced fidelity`)
        // non-blocking: the preview and BOTH overlays still render normally
        assert.ok(harness.overlay('overlay-evidence-region'), `${backend} must still render the region`)
        assert.equal(harness.panel().querySelectorAll('img').length, 2)
      } else {
        assert.doesNotMatch(harness.panel().textContent, notice, 'PowerPoint must not show a fallback warning')
      }
    }
  } finally {
    await harness.cleanup()
  }
})

/* ------------------------------------------------------------------ §17 compact overlay toggle */
test('the compact overlay toggle switches geometry without unmounting the review', async () => {
  const harness = await mountLearningHarness()
  try {
    await harness.openImageReview()
    await harness.setLearning([candidate({
      id: 'image-1', slide: 5, pictureId: '42', itemIndex: 1,
      targetPct: { x: 50, y: 35, w: 12, h: 14 },
      evidencePct: { x: 47, y: 31, w: 33, h: 24 },
      backend: 'powerpoint', faithful: true,
      evidenceMeta: { itemId: 'item-two', itemIndex: 1, itemHeading: 'Mục hai', pictureCount: 3 },
    })])
    await harness.waitFor(() => harness.overlay('overlay-evidence-region'), 'the default overlay set')

    // default = both, so the FINAL Excel region is obvious straight away (§17)
    assert.ok(harness.overlay('overlay-evidence-region') && harness.overlay('overlay-picture-candidate'))
    const group = harness.container.querySelector('[aria-label="Chọn khung hiển thị trên slide"]')
    assert.ok(group, 'the compact toggle group must render')
    assert.deepEqual([...group.querySelectorAll('button')].map((b) => b.textContent.trim()),
      ['Vùng xuất Excel', 'Ảnh đang đánh giá', 'Cả hai'])

    await harness.click(harness.byText('button', 'Ảnh đang đánh giá'))
    assert.equal(harness.overlay('overlay-evidence-region'), null, 'evidence overlay hidden')
    assert.ok(harness.overlay('overlay-picture-candidate'), 'picture overlay still shown')

    await harness.click(harness.byText('button', 'Vùng xuất Excel'))
    assert.ok(harness.overlay('overlay-evidence-region'), 'evidence overlay shown alone')
    assert.equal(harness.overlay('overlay-picture-candidate'), null)

    await harness.click(harness.byText('button', 'Cả hai'))
    assert.ok(harness.overlay('overlay-evidence-region') && harness.overlay('overlay-picture-candidate'))

    // §44: the component is still the SAME mounted review — hook order survived every toggle
    assert.equal(harness.panel().querySelector('h2').textContent, 'Slide 5')
  } finally {
    await harness.cleanup()
  }
})

/* ------------------------------------------------------------------ §20/§31 item scoping in the UI */
test('two pictures of one item share its region and a picture with no region shows none', async () => {
  const harness = await mountLearningHarness()
  try {
    await harness.openImageReview()
    const regionOfItem2 = { x: 47, y: 31, w: 33, h: 24 }
    const shared = {
      slide: 5, itemIndex: 1, backend: 'powerpoint', faithful: true,
      evidencePct: regionOfItem2,
      evidenceMeta: { itemId: 'item-two', itemIndex: 1, itemHeading: 'Mục hai', pictureCount: 2 },
    }
    const first = candidate({ ...shared, id: 'image-a', pictureId: '42', targetPct: { x: 50, y: 35, w: 12, h: 14 } })
    const second = candidate({ ...shared, id: 'image-b', pictureId: '43', targetPct: { x: 64, y: 35, w: 12, h: 14 } })
    // a candidate no eligible After region owns: it must NOT borrow Mục #2's region (§20/§31)
    const orphan = candidate({
      ...shared, id: 'image-c', slide: 6, pictureId: '99', itemIndex: -1,
      evidencePct: null,
      evidenceMeta: { itemId: '', itemIndex: -1, itemHeading: '', pictureCount: 0 },
      targetPct: { x: 10, y: 10, w: 8, h: 8 },
    })

    await harness.setLearning([first, second, orphan])
    await harness.waitFor(() => harness.overlay('overlay-evidence-region'), 'the first candidate region')

    const regionBoxA = harness.boxOf(harness.overlay('overlay-evidence-region'))
    const pictureBoxA = harness.boxOf(harness.overlay('overlay-picture-candidate'))

    const listButtons = () => [...harness.container.querySelectorAll('section[aria-label="Đối tượng ảnh"] button')]
    assert.equal(listButtons().length, 3, 'all three candidates are listed')
    await harness.click(listButtons()[1])                       // the SECOND picture of Mục #2
    await harness.waitFor(
      () => harness.boxOf(harness.overlay('overlay-picture-candidate')).left === 64,
      'the second candidate of the same item',
    )
    // §56-L: switching between two pictures of Mục #2 KEEPS Mục #2's final region
    assert.deepEqual(harness.boxOf(harness.overlay('overlay-evidence-region')), regionBoxA)
    assert.notDeepEqual(harness.boxOf(harness.overlay('overlay-picture-candidate')), pictureBoxA)
    assert.match(harness.overlay('overlay-evidence-region').textContent, /Mục #2 · Vùng xuất Excel/)

    await harness.click(listButtons()[2])                       // the picture no region owns
    await harness.waitFor(() => harness.panel().querySelector('h2').textContent === 'Slide 6', 'the orphan candidate')
    assert.equal(harness.overlay('overlay-evidence-region'), null, 'no region overlay for an unowned picture')
    assert.ok(harness.overlay('overlay-picture-candidate'), 'the picture candidate overlay remains')
    assert.match(harness.container.textContent, /KHÔNG nằm trong vùng Sau cải tiến nào sẽ được xuất Excel/)
    assert.match(harness.panel().textContent, /không thuộc vùng Sau cải tiến nào sẽ xuất Excel/)
  } finally {
    await harness.cleanup()
  }
})

/* ------------------------------------------------------------------ §44 hook order with no candidate */
test('hook order stays stable when the review has no candidate (React #310 preserved)', async () => {
  const harness = await mountLearningHarness()
  try {
    await harness.openImageReview()
    await harness.setLearning([])
    await harness.waitFor(
      () => /Python backend chưa trả về ảnh ứng viên cần kiểm tra\./.test(harness.container.textContent),
      'the empty-state render',
    )
    // toggling overlay mode while there is no candidate must not crash the hook order
    const group = harness.container.querySelector('[aria-label="Chọn khung hiển thị trên slide"]')
    assert.equal(group, null, 'the overlay toggle belongs to the preview panel only')
    await harness.setLearning([candidate({
      id: 'image-1', slide: 5, pictureId: '42', itemIndex: 0,
      targetPct: { x: 20, y: 25, w: 30, h: 35 },
      evidencePct: { x: 18, y: 22, w: 40, h: 45 },
      backend: 'builtin', faithful: false,
      evidenceMeta: { itemId: 'item-one', itemIndex: 0, itemHeading: 'Mục một', pictureCount: 1 },
    })])
    await harness.waitFor(() => harness.overlay('overlay-evidence-region'), 'recovery after the empty state')
  } finally {
    await harness.cleanup()
  }
})

/* ------------------------------------------------------------------ §50 visible build identity */
test('the header shows version AND build so a stale Windows checkout is obvious', async () => {
  const harness = await mountLearningHarness()
  try {
    const badge = harness.container.querySelector('[data-testid="app-version-badge"]')
    assert.ok(badge, 'the version badge must render')
    assert.equal(badge.textContent, `v${APP_VERSION} · Build ${APP_BUILD}`)
    assert.match(badge.textContent, /v1\.3\.3 · Build 016/)
    // no duplicated version text in the product title next to it
    assert.equal(harness.container.querySelector('h1').textContent, 'Report Extractor')
  } finally {
    await harness.cleanup()
  }
})

/* ------------------------------------------------------------------ §51 frontend/backend agreement */
test('the frontend package metadata agrees with the authoritative backend version', async () => {
  const { readFile } = await import('node:fs/promises')
  const { join } = await import('node:path')
  const pkg = JSON.parse(await readFile(join(frontendRoot, 'package.json'), 'utf8'))
  const lock = JSON.parse(await readFile(join(frontendRoot, 'package-lock.json'), 'utf8'))
  assert.equal(pkg.version, APP_VERSION)
  assert.equal(lock.version, APP_VERSION)
  assert.equal(lock.packages[''].version, APP_VERSION)

  const pythonInit = await readFile(join(frontendRoot, '..', 'app', '__init__.py'), 'utf8')
  const version = pythonInit.match(/__version__\s*=\s*"([^"]+)"/)[1]
  const buildNumber = Number(pythonInit.match(/BUILD_NUMBER\s*=\s*(\d+)/)[1])
  // This is the §51 regression guard: it FAILS if the two sides drift (backend 1.3.3 / frontend 1.3.2).
  assert.equal(pkg.version, version, 'frontend/package.json must match app/__init__.py __version__')
  assert.equal(APP_VERSION, version)
  assert.equal(APP_BUILD, String(buildNumber).padStart(3, '0'))

  const indexHtml = await readFile(join(frontendRoot, 'index.html'), 'utf8')
  assert.match(indexHtml, new RegExp(`Report Extractor — ${version} — Build ${APP_BUILD}`))

  // the header must read the bridge value, never hardcode its own copy of the version (§50/§51)
  const header = await readFile(join(frontendRoot, 'src', 'components', 'AppHeader.tsx'), 'utf8')
  assert.doesNotMatch(header, new RegExp(version.replace(/\./g, '\\.')))
  assert.doesNotMatch(header, new RegExp(`Build ${APP_BUILD}`))
  assert.match(header, /appInfo\.version/)
  assert.match(header, /appInfo\.build/)
})


