import assert from 'node:assert/strict'
import { test } from 'node:test'
import { JSDOM } from 'jsdom'
import { createServer } from 'vite'
import * as React from 'react'
import { fileURLToPath } from 'node:url'

const frontendRoot = fileURLToPath(new URL('..', import.meta.url))
const delay = (ms) => new Promise((resolve) => globalThis.setTimeout(resolve, ms))

function emptyLearningState(imageCount = 0, contentCount = 0) {
  return {
    images: [],
    contents: [],
    candidateStatus: {
      image: { state: 'ready', reason: imageCount ? 'available' : 'empty', count: imageCount,
        message: imageCount ? `Đã tải ${imageCount} ảnh ứng viên từ lượt xử lý hiện tại.` : 'Lượt xử lý hiện tại không có ảnh ứng viên cần kiểm tra.' },
      content: { state: 'ready', reason: contentCount ? 'available' : 'empty', count: contentCount,
        message: contentCount ? `Đã tải ${contentCount} khối nội dung từ lượt xử lý hiện tại.` : 'Lượt xử lý hiện tại không có khối nội dung cần kiểm tra.' },
    },
    counts: { image: { total: 0, labeled: 0 }, content: { total: 0, labeled: 0 } },
    modelStatus: { image: 'Ready', content: 'Ready' },
    excelPending: 0,
    excelLastResult: { kind: 'none', message: '' },
    training: false,
  }
}

function contentCandidate(id = 'content-1') {
  return {
    id,
    sourceFile: 'improvement-report.pptx',
    managementNumber: 'TNP-001',
    slide: 3,
    blockId: 'shape-17',
    text: 'Nội dung cải tiến nguyên văn từ PowerPoint',
    decision: 'include',
    confidenceBand: 'high',
    confidence: 0.96,
    evidence: ['Thuộc vùng cải tiến trong sản xuất'],
    section: 'improvement',
    nearestTitle: 'Cải tiến trong sản xuất',
    userLabel: 'UNLABELED',
    note: '',
  }
}

function imageCandidate(id, slide, itemIndex, slidePreview) {
  return {
    id,
    sourceFile: 'improvement-report.pptx',
    managementNumber: 'TNP-001',
    slide,
    pictureId: `picture-${slide}`,
    src: 'data:image/png;base64,AA==',
    bounds: { x: 20, y: 25, w: 30, h: 35 },
    slideWidth: 9144000,
    slideHeight: 5143500,
    targetBboxPct: { x: 20, y: 25, w: 30, h: 35 },
    slidePreview,
    slidePreviewWidth: 1600,
    slidePreviewHeight: 900,
    itemId: `item-${slide}`,
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

async function waitFor(predicate, description, timeoutMs = 2_000) {
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    if (predicate()) return
    await React.act(async () => { await delay(10) })
  }
  assert.fail(`Timed out waiting for ${description}`)
}

function bridgeEnvelope(data) {
  return { ok: true, data }
}

test('ImageReview keeps hook order through StrictMode disconnect, readiness, candidate loading, switching and preview', async () => {
  const dom = new JSDOM('<!doctype html><html><body><div id="root"></div></body></html>', {
    url: 'http://localhost/',
  })
  const globalNames = [
    'window', 'document', 'navigator', 'HTMLElement', 'Node', 'Event', 'MouseEvent',
    'ResizeObserver', 'IS_REACT_ACT_ENVIRONMENT',
  ]
  const priorGlobals = new Map(globalNames.map((name) => [name, Object.getOwnPropertyDescriptor(globalThis, name)]))
  const defineGlobal = (name, value) => Object.defineProperty(globalThis, name, {
    configurable: true,
    writable: true,
    value,
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

  let server
  let root
  try {
    server = await createServer({
      configFile: `${frontendRoot}/vite.config.ts`,
      root: frontendRoot,
      appType: 'custom',
      server: { middlewareMode: true, hmr: false },
    })
    const [{ StoreProvider, useStore }, { LearningTab }] = await Promise.all([
      server.ssrLoadModule('/src/state/store.tsx'),
      server.ssrLoadModule('/src/tabs/LearningTab.tsx'),
    ])
    const { createRoot } = await import('react-dom/client')

    const config = {
      paths: { reportFolder: '', template: '', output: '' },
      period: { mode: 'auto', month: '', year: '', from: '', to: '' },
      forceReprocess: false,
      ollama: { host: '', port: '11434', model: '', checked: 'unchecked', message: '', models: [] },
      update: { path: '', autoCheck: true },
    }
    const dashboard = {
      app: { name: 'Report Extractor', title: 'Report Extractor v1.3.4', version: '1.3.4', build: '017', buildNumber: 17 },
      config,
      scan: { scanned: false, message: 'No reports selected.' },
      reports: [],
      job: {
        status: 'idle', queue: [], index: 0, doneCount: 0, stage: 'waiting', percent: 0,
        currentFile: '', elapsedSec: 0, remainSec: null, startedAt: null, finishedAt: null, hasSamples: false,
      },
      logs: [],
      ollama: {
        host: '', port: '11434', model: '', checked: 'unchecked', message: '', models: [], aiStatus: '',
        discovering: false, discoveryMessage: '', discoveryResults: [], discoveryChecked: 0,
        discoveryTotal: 0, serverApplyRunning: false, serverApplyMessage: '',
      },
      update: { path: '', autoCheck: true, status: 'idle', message: '', progress: 0, progressStage: '' },
      diagnostics: { running: false, rows: [], error: '' },
    }

    let learningResponse = emptyLearningState()
    let learningFailure = null
    let learningCalls = 0
    // Start disconnected, just as the desktop UI does before pywebview injects its API.

    function Harness() {
      const store = useStore()
      return React.createElement(
        React.Fragment,
        null,
        React.createElement('output', { 'data-testid': 'bridge-state' }, store.connected ? 'ready' : 'disconnected'),
        React.createElement('button', {
          type: 'button',
          'data-testid': 'enter-learning',
          onClick: () => { store.setTab('learning') },
        }, 'Enter Learning'),
        React.createElement('button', {
          type: 'button',
          'data-testid': 'refresh-learning',
          onClick: () => { void store.refreshLearning() },
        }, 'Refresh candidates'),
        React.createElement(LearningTab),
      )
    }

    root = createRoot(dom.window.document.getElementById('root'))
    await React.act(async () => {
      root.render(React.createElement(React.StrictMode, null,
        React.createElement(StoreProvider, null, React.createElement(Harness))))
    })

    const container = dom.window.document.getElementById('root')
    const imageReviewMenu = [...container.querySelectorAll('aside nav button')]
      .find((button) => button.textContent.trim() === 'Kiểm tra ảnh cải tiến')
    assert.ok(imageReviewMenu, 'image-review menu should render')
    await React.act(async () => {
      imageReviewMenu.dispatchEvent(new dom.window.MouseEvent('click', { bubbles: true }))
    })
    assert.match(container.textContent, /Đang chờ kết nối Python backend\./)
    assert.equal(container.querySelector('[data-testid="bridge-state"]').textContent, 'disconnected')

    // Establish the native bridge after the empty/disconnected ImageReview has already rendered.
    dom.window.pywebview = {
      api: {
        ping: async () => bridgeEnvelope('Python bridge OK'),
        get_app_version: async () => bridgeEnvelope({
          name: 'Report Extractor', title: 'Report Extractor v1.3.4', version: '1.3.4', build: '017',
        }),
        get_current_config: async () => bridgeEnvelope(config),
        get_dashboard_state: async () => bridgeEnvelope(dashboard),
        get_learning_state: async () => {
          learningCalls += 1
          return learningFailure
            ? { ok: false, error: { code: 'LEARNING_FAILED', message: learningFailure } }
            : bridgeEnvelope(learningResponse)
        },
      },
    }
    await React.act(async () => {
      dom.window.dispatchEvent(new dom.window.Event('pywebviewready'))
    })
    await waitFor(
      () => container.querySelector('[data-testid="bridge-state"]')?.textContent === 'ready',
      'the pywebview startup handshake',
    )

    // Entering Learning and the manual control each perform a real backend read. A successful zero-candidate
    // snapshot is visible as valid empty state, never as "Python did not return a field".
    await React.act(async () => {
      container.querySelector('[data-testid="enter-learning"]').dispatchEvent(
        new dom.window.MouseEvent('click', { bubbles: true }),
      )
      await delay(0)
    })
    await waitFor(() => learningCalls > 0, 'the Learning-entry refresh')
    const callsAfterEntry = learningCalls
    await React.act(async () => {
      container.querySelector('[data-testid="refresh-learning"]').dispatchEvent(
        new dom.window.MouseEvent('click', { bubbles: true }),
      )
      await delay(0)
    })
    await waitFor(() => learningCalls > callsAfterEntry, 'the manual learning-state refresh')
    assert.match(container.textContent, /Lượt xử lý hiện tại không có ảnh ứng viên cần kiểm tra\./)

    const previewOne = 'data:image/svg+xml;base64,PHN2ZyBpZD0ib25lIi8+'
    const previewTwo = 'data:image/svg+xml;base64,PHN2ZyBpZD0idHdvIi8+'
    learningResponse = {
      ...emptyLearningState(2, 1),
      images: [
        imageCandidate('image-1', 1, 0, previewOne),
        imageCandidate('image-2', 2, 1, previewTwo),
      ],
      contents: [contentCandidate()],
      counts: { image: { total: 2, labeled: 0 }, content: { total: 1, labeled: 0 } },
    }
    await React.act(async () => {
      container.querySelector('[data-testid="refresh-learning"]').dispatchEvent(
        new dom.window.MouseEvent('click', { bubbles: true }),
      )
      await delay(0)
    })

    await waitFor(() => container.querySelector('section[aria-label="Slide ngữ cảnh"] h2')?.textContent === 'Slide 1',
    'the populated candidate and first slide preview')

    let previewPanel = container.querySelector('section[aria-label="Slide ngữ cảnh"]')
    // PROMPT-027R §18/§19: ONE unmodified slide bitmap (no blurred duplicate, no clipped copy)
    assert.equal(previewPanel.querySelectorAll('img').length, 1, 'the unmodified slide bitmap should render once')
    assert.deepEqual(
      [...previewPanel.querySelectorAll('img')].map((image) => image.getAttribute('src')),
      [previewOne],
    )

    const secondCandidate = [...container.querySelectorAll('section[aria-label="Đối tượng ảnh"] button')]
      .find((button) => button.textContent.includes('Slide 2'))
    assert.ok(secondCandidate, 'the second candidate should appear in the review list')
    await React.act(async () => {
      secondCandidate.dispatchEvent(new dom.window.MouseEvent('click', { bubbles: true }))
    })
    await waitFor(
      () => container.querySelector('section[aria-label="Slide ngữ cảnh"] h2')?.textContent === 'Slide 2',
      'switching to the second candidate',
    )
    previewPanel = container.querySelector('section[aria-label="Slide ngữ cảnh"]')
    assert.deepEqual(
      [...previewPanel.querySelectorAll('img')].map((image) => image.getAttribute('src')),
      [previewTwo],
      'switching candidates should update the rendered slide preview',
    )

    const previousCandidate = [...container.querySelectorAll('section[aria-label="Thông tin và nhãn xác nhận"] button')]
      .find((button) => button.textContent.trim().includes('Mục trước'))
    assert.ok(previousCandidate, 'the previous-candidate control should render')
    await React.act(async () => {
      previousCandidate.dispatchEvent(new dom.window.MouseEvent('click', { bubbles: true }))
    })
    await waitFor(
      () => container.querySelector('section[aria-label="Slide ngữ cảnh"] h2')?.textContent === 'Slide 1',
      'switching back to the first candidate',
    )

    const openLearningSection = async (label) => {
      const button = [...container.querySelectorAll('aside nav button')]
        .find((entry) => entry.textContent.trim() === label)
      assert.ok(button, `${label} section button should render`)
      await React.act(async () => {
        button.dispatchEvent(new dom.window.MouseEvent('click', { bubbles: true }))
      })
    }
    const refresh = async () => {
      const before = learningCalls
      await React.act(async () => {
        container.querySelector('[data-testid="refresh-learning"]').dispatchEvent(
          new dom.window.MouseEvent('click', { bubbles: true }),
        )
        await delay(0)
      })
      await waitFor(() => learningCalls > before, 'a fresh manual Learning bridge call')
    }

    // Content data is independently visible and remains verbatim.
    await openLearningSection('Kiểm tra nội dung cải tiến')
    await waitFor(() => container.textContent.includes('Nội dung cải tiến nguyên văn từ PowerPoint'),
      'the populated content candidate')

    // Successful zero-candidate image/content snapshots have explicit, truthful feedback.
    learningResponse = emptyLearningState()
    await refresh()
    await waitFor(() => container.querySelector('[data-testid="learning-content-feedback"]')
      ?.textContent.includes('không có khối nội dung cần kiểm tra'), 'valid empty content feedback')
    await openLearningSection('Kiểm tra ảnh cải tiến')
    assert.match(container.querySelector('[data-testid="learning-image-feedback"]').textContent,
      /không có ảnh ứng viên cần kiểm tra/)

    // A successful envelope missing a required field is a contract failure, not valid empty.
    learningResponse = emptyLearningState()
    delete learningResponse.contents
    await refresh()
    await waitFor(() => container.querySelector('[data-testid="learning-image-feedback"]')
      ?.textContent.includes('Không thể tải trạng thái học'), 'missing-field image failure feedback')
    await openLearningSection('Kiểm tra nội dung cải tiến')
    assert.match(container.querySelector('[data-testid="learning-content-feedback"]').textContent,
      /Không thể tải trạng thái học/)

    // A backend rejection is visible in both paths, and the same manual control can recover with a fresh call.
    learningResponse = emptyLearningState()
    learningFailure = 'Không đọc được trạng thái học hiện tại.'
    await refresh()
    await waitFor(() => container.querySelector('[data-testid="learning-content-feedback"]')
      ?.textContent.includes('Không đọc được trạng thái học hiện tại'), 'backend content failure feedback')
    await openLearningSection('Kiểm tra ảnh cải tiến')
    assert.match(container.querySelector('[data-testid="learning-image-feedback"]').textContent,
      /Không đọc được trạng thái học hiện tại/)
    learningFailure = null
    await refresh()
    await waitFor(() => container.querySelector('[data-testid="learning-image-feedback"]')
      ?.textContent.includes('không có ảnh ứng viên cần kiểm tra'), 'manual refresh recovery')
  } finally {
    if (root) {
      await React.act(async () => { root.unmount() })
    }
    if (server) await server.close()
    dom.window.close()
    for (const [name, descriptor] of priorGlobals) {
      if (descriptor) Object.defineProperty(globalThis, name, descriptor)
      else delete globalThis[name]
    }
  }
})
