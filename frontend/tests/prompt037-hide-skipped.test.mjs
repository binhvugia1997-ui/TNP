// PROMPT-037 regression: skipped reports (status `skipped`, “Bỏ qua — đã cập nhật”) are hidden from
// “Danh sách báo cáo” and from its counters, selection and active row. The DTO keeps the row; the
// presentation rule is the only filter. Store-level tests drive the real store through a mocked
// pywebview bridge; the rendered ReportTable and StatsCard are checked against the store.
import assert from 'node:assert/strict'
import { after, before, test } from 'node:test'
import { JSDOM } from 'jsdom'
import { createServer } from 'vite'
import * as React from 'react'
import { fileURLToPath } from 'node:url'

const frontendRoot = fileURLToPath(new URL('..', import.meta.url))
let server
let StoreProvider
let useStore
let resolveActiveReportId
let presentReports
let isVisibleReport
let HIDDEN_REPORT_STATUSES
let BUCKET_OF
let STATUS_LABEL
let ReportTable
let StatsCard

before(async () => {
  server = await createServer({
    configFile: `${frontendRoot}/vite.config.ts`,
    root: frontendRoot,
    appType: 'custom',
    server: { middlewareMode: true, hmr: false },
  })
  ;({ StoreProvider, useStore, resolveActiveReportId } = await server.ssrLoadModule('/src/state/store.tsx'))
  ;({ presentReports, isVisibleReport, HIDDEN_REPORT_STATUSES, BUCKET_OF, STATUS_LABEL } =
    await server.ssrLoadModule('/src/types.ts'))
  ;({ ReportTable } = await server.ssrLoadModule('/src/components/report/ReportTable.tsx'))
  ;({ StatsCard } = await server.ssrLoadModule('/src/components/report/LogAndStats.tsx'))
})

after(async () => {
  await server?.close()
})

const delay = (ms) => new Promise((resolve) => globalThis.setTimeout(resolve, ms))
const envelope = (data) => ({ ok: true, data })

// Bridge methods the hide flows are allowed to use. Anything else (exclude/restore/delete/…) is a failure.
const READ_OR_SCAN_CALLS = new Set(['ping', 'get_app_version', 'get_current_config', 'get_dashboard_state',
  'scan_reports', 'get_report_details'])

function rep(id, status, fileName = `report-${id}.pptx`) {
  return {
    id: String(id), stt: Number(id) + 1, managementNumber: `2610101${String(id).padStart(4, '0')}-VOC`,
    fileName, path: '', occurrenceDate: '10/10/2026', vendor: '', manualFields: {}, slides: 0, status,
    shortResult: '', results: { causes: [], countermeasures: [], temporaryRemoved: false, afterImages: [] },
  }
}

function dashboard(scan, reports = []) {
  const config = {
    paths: { reportFolder: 'reports', template: '', output: '' },
    period: { mode: 'all', month: '', year: '', from: '', to: '' },
    forceReprocess: false,
    ollama: { host: '', port: '11434', model: '', checked: 'unchecked', message: '', models: [] },
    update: { path: '', autoCheck: false },
  }
  return {
    app: { name: 'Report Extractor', title: 'Report Extractor', version: '1.3.6', build: '019', buildNumber: 19 },
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

async function waitFor(predicate, description, timeoutMs = 4_000) {
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    if (predicate()) return
    await React.act(async () => { await delay(15) })
  }
  assert.fail(`Timed out waiting for ${description}`)
}

// Mounts the real StoreProvider with ReportTable + StatsCard and a recording mock bridge.
async function mount(initialReports, initialScan = { scanned: true, message: 'ready', generation: 1, revision: 1, inProgress: false }) {
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

  const state = {
    dashboard: dashboard(initialScan, initialReports),
    scanReply: null,
    calls: [],
    store: null,
    dashboardReads: 0,
  }
  const handlers = {
    ping: async () => envelope('ok'),
    get_app_version: async () => envelope(state.dashboard.app),
    get_current_config: async () => envelope(state.dashboard.config),
    get_dashboard_state: async () => { state.dashboardReads += 1; return envelope(state.dashboard) },
    scan_reports: async () => envelope(state.scanReply ?? state.dashboard),
    get_report_details: async () => envelope({ slides: 0, causes: [], countermeasures: [], afterImages: [] }),
  }
  // Any method without a handler is recorded and fails loudly in assertions (no silent success).
  dom.window.pywebview = { api: new Proxy({}, {
    get(_target, method) {
      return async (...args) => {
        state.calls.push(String(method))
        if (handlers[method]) return handlers[method](...args)
        return envelope(state.dashboard)
      }
    },
  }) }

  function Harness() {
    state.store = useStore()
    return React.createElement(React.Fragment, null,
      React.createElement(ReportTable),
      React.createElement(StatsCard))
  }
  const { createRoot } = await import('react-dom/client')
  const root = createRoot(dom.window.document.getElementById('root'))
  await React.act(async () => { root.render(React.createElement(StoreProvider, null, React.createElement(Harness))) })
  await waitFor(() => state.store?.connected === true, 'bridge startup')

  let revision = initialScan.revision ?? 1
  const setDashboard = (reports, scanPatch = {}) => {
    revision += 1
    state.dashboard = dashboard({ scanned: true, message: 'ready', generation: 1, revision, inProgress: false,
      ...scanPatch }, reports)
  }
  return {
    state,
    store: () => state.store,
    ids: () => state.store.reports.map((report) => report.id),
    visibleIds: () => state.store.filteredReports.map((report) => report.id),
    setDashboard,
    document: () => dom.window.document,
    MouseEvent: dom.window.MouseEvent,
    rows: () => [...dom.window.document.querySelectorAll('tbody tr')],
    text: () => dom.window.document.body.textContent,
    act: (fn) => React.act(async () => { await fn() }),
    waitFor,
    cleanup: async () => {
      await React.act(async () => { root.unmount() })
      dom.window.close()
      for (const [name, descriptor] of prior) {
        if (descriptor) Object.defineProperty(globalThis, name, descriptor)
        else delete globalThis[name]
      }
    },
  }
}

const sameList = (actual, expected) => assert.deepEqual([...actual], [...expected])
const mutationCalls = (calls) => calls.filter((name) => !READ_OR_SCAN_CALLS.has(name))

// ---------------------------------------------------------------- pure visibility rule

test('visibility rule: waiting, completed, needs_review and error reports stay visible', () => {
  for (const status of ['waiting', 'new_row', 'processing', 'completed', 'needs_review', 'error', 'cancelled',
    'excluded', 'outside_period', 'source_duplicate', 'fast_skip']) {
    assert.equal(isVisibleReport({ status }), true, `${status} must stay visible`)
  }
})

test('visibility rule: a skipped report is absent from the visible list and is the only hidden status', () => {
  assert.equal(isVisibleReport({ status: 'skipped' }), false)
  assert.deepEqual([...HIDDEN_REPORT_STATUSES], ['skipped'])
})

test('mixed list: only skipped reports are removed and STT is renumbered over the visible rows', () => {
  const input = [rep(0, 'waiting'), rep(1, 'skipped'), rep(2, 'completed'), rep(3, 'skipped'), rep(4, 'needs_review'),
    rep(5, 'error')]
  const out = presentReports(input)
  sameList(out.map((r) => r.id), ['0', '2', '4', '5'])
  sameList(out.map((r) => r.stt), [1, 2, 3, 4])
  // Unchanged rows keep their identity; only STT-changed rows are copied.
  assert.equal(out[0], input[0])
  assert.notEqual(out[1], input[2])
  assert.equal(out[1].stt, 2)
})

test('1000+ mixed reports: exact visible count and linear, non-pathological cost', () => {
  const total = 12_000
  const input = Array.from({ length: total }, (_, i) => rep(i, i % 4 === 0 ? 'skipped'
    : ['waiting', 'completed', 'needs_review', 'error'][i % 3]))
  const expected = input.filter((r) => r.status !== 'skipped').length
  const started = performance.now()
  const out = presentReports(input)
  const elapsed = performance.now() - started
  assert.equal(out.length, expected)
  assert.equal(out.length, total - total / 4)
  assert.ok(out.every((r) => r.status !== 'skipped'))
  assert.deepEqual(out.map((r) => r.stt), Array.from({ length: out.length }, (_, i) => i + 1))
  assert.ok(elapsed < 500, `presentReports took ${elapsed.toFixed(1)} ms for ${total} rows`)
})

test('resolveActiveReportId keeps a visible row, advances to the next visible one, then the previous, else clears', () => {
  const previous = ['a', 'b', 'c', 'd'].map((id) => rep(id, 'waiting', `${id}.pptx`))
  const without = (ids) => previous.filter((r) => ids.includes(r.id))
  assert.equal(resolveActiveReportId(previous, without(['a', 'b', 'c', 'd']), 'b'), 'b')
  assert.equal(resolveActiveReportId(previous, without(['a', 'c', 'd']), 'b'), 'c', 'next visible after b')
  assert.equal(resolveActiveReportId(previous, without(['a', 'b', 'c']), 'd'), 'c', 'previous visible when last hidden')
  assert.equal(resolveActiveReportId(previous, [], 'b'), '', 'empty list clears the active row')
  assert.equal(resolveActiveReportId(previous, without(['c', 'd']), ''), 'c', 'no active row → first visible')
})

test('skip classification semantics unchanged in the frontend contract', () => {
  assert.equal(BUCKET_OF.skipped, 'skipped')
  assert.equal(STATUS_LABEL.skipped, 'Bỏ qua — đã cập nhật')
})

// ---------------------------------------------------------------- store + rendered list

test('a full dashboard hides skipped rows; waiting/completed/needs_review/error stay; header and totals match the list', async () => {
  const m = await mount([rep(0, 'waiting'), rep(1, 'skipped'), rep(2, 'completed'), rep(3, 'needs_review'),
    rep(4, 'error')])
  try {
    sameList(m.ids(), ['0', '2', '3', '4'])
    assert.equal(m.store().stats.total, 4)
    assert.equal(m.store().stats.skipped, 0)
    assert.equal(m.rows().length, 4)
    assert.match(m.text(), /4\/4 file/)
    assert.match(m.text(), /Tổng số báo cáo: 4/)
    assert.doesNotMatch(m.text(), /report-1\.pptx/)
  } finally {
    await m.cleanup()
  }
})

test('dashboard refresh cannot reinsert a skipped report, and a visible row turned skipped leaves on the next poll', async () => {
  const m = await mount([rep(0, 'waiting'), rep(1, 'completed'), rep(2, 'waiting')])
  try {
    sameList(m.ids(), ['0', '1', '2'])
    m.setDashboard([rep(0, 'waiting'), rep(1, 'skipped'), rep(2, 'waiting')])
    await waitFor(() => m.ids().join(',') === '0,2', 'skipped row removed by poll')
    const reads = m.state.dashboardReads
    await m.act(() => delay(1_600)) // several polls keep returning the skipped row
    assert.ok(m.state.dashboardReads >= reads + 2, 'dashboard polling continued')
    sameList(m.ids(), ['0', '2'])
    assert.equal(m.store().stats.total, 2)
    assert.deepEqual(mutationCalls(m.state.calls), [], 'hiding issues no mutation or delete bridge call')
  } finally {
    await m.cleanup()
  }
})

test('incremental scan deltas cannot reinsert a skipped report, and a visible→skipped delta removes it', async () => {
  const m = await mount([])
  try {
    let resolveScan
    m.state.scanReply = new Promise((resolve) => { resolveScan = resolve })
    await m.act(() => { void m.store().scan() })
    await waitFor(() => m.store().scanBusy === true, 'scan started')

    // leaf 1: a waiting row and a skipped row arrive together
    m.state.dashboard = dashboard({ scanned: true, message: 'leaf 1', generation: 2, revision: 1, inProgress: true,
      incremental: true, completedLeaves: 1, totalLeaves: 3 }, [rep(0, 'waiting'), rep(1, 'skipped')])
    await waitFor(() => m.ids().join(',') === '0', 'leaf 1 visible rows')

    // leaf 2: the skipped row is delivered again plus a completed row
    m.state.dashboard = dashboard({ scanned: true, message: 'leaf 2', generation: 2, revision: 2, inProgress: true,
      incremental: true, completedLeaves: 2, totalLeaves: 3 }, [rep(1, 'skipped'), rep(2, 'completed')])
    await waitFor(() => m.ids().join(',') === '0,2', 'leaf 2 visible rows')
    await m.act(() => delay(900))
    sameList(m.ids(), ['0', '2'])

    // leaf 3: row 0 transitions visible → skipped inside a delta
    m.state.dashboard = dashboard({ scanned: true, message: 'leaf 3', generation: 2, revision: 3, inProgress: true,
      incremental: true, completedLeaves: 3, totalLeaves: 3 }, [rep(0, 'skipped')])
    await waitFor(() => m.ids().join(',') === '2', 'visible→skipped delta removes row 0')
    assert.equal(m.store().stats.total, 1)

    resolveScan(dashboard({ scanned: true, message: 'done', generation: 2, revision: 4, inProgress: false },
      [rep(0, 'skipped'), rep(1, 'skipped'), rep(2, 'completed')]))
    await waitFor(() => m.store().scanBusy === false, 'final scan response')
    sameList(m.ids(), ['2'])
    assert.deepEqual(mutationCalls(m.state.calls), [], 'incremental hide issues no mutation or delete bridge call')
  } finally {
    await m.cleanup()
  }
})

test('a rescan with the same skipped classification stays hidden', async () => {
  const m = await mount([rep(0, 'waiting'), rep(1, 'skipped'), rep(2, 'completed')])
  try {
    sameList(m.ids(), ['0', '2'])
    m.state.scanReply = dashboard({ scanned: true, message: 'rescan', generation: 2, revision: 1, inProgress: false },
      [rep(0, 'waiting'), rep(1, 'skipped'), rep(2, 'completed')])
    await m.act(() => m.store().scan())
    sameList(m.ids(), ['0', '2'])
    assert.equal(m.store().stats.total, 2)
    assert.ok(m.state.calls.includes('scan_reports'))
    await m.act(() => delay(900))
    sameList(m.ids(), ['0', '2'])
    assert.deepEqual(mutationCalls(m.state.calls), [], 'rescan hide issues no mutation or delete bridge call')
  } finally {
    await m.cleanup()
  }
})

test('a selected report that becomes skipped is removed from selectedIds', async () => {
  const m = await mount([rep(0, 'waiting'), rep(1, 'waiting'), rep(2, 'completed')])
  try {
    await m.act(() => { m.store().toggleSelected('0'); m.store().toggleSelected('2') })
    sameList(m.store().selectedIds, ['0', '2'])
    m.setDashboard([rep(0, 'waiting'), rep(1, 'waiting'), rep(2, 'skipped')])
    await waitFor(() => m.store().selectedIds.length === 1, 'skipped selection pruned')
    sameList(m.store().selectedIds, ['0'])
  } finally {
    await m.cleanup()
  }
})

test('an active row that becomes skipped advances to the next visible row; the last one falls back to the previous; empty clears', async () => {
  const m = await mount([rep(0, 'waiting'), rep(1, 'waiting'), rep(2, 'waiting'), rep(3, 'waiting')])
  try {
    await m.act(() => m.store().setActiveId('1'))
    assert.equal(m.store().activeId, '1')
    m.setDashboard([rep(0, 'waiting'), rep(1, 'skipped'), rep(2, 'waiting'), rep(3, 'waiting')])
    await waitFor(() => m.store().activeId === '2', 'active advanced to next visible row')
    assert.equal(m.store().activeReport?.id, '2')

    await m.act(() => m.store().setActiveId('3'))
    m.setDashboard([rep(0, 'waiting'), rep(1, 'skipped'), rep(2, 'waiting'), rep(3, 'skipped')])
    await waitFor(() => m.store().activeId === '2', 'active fell back to previous visible row')

    m.setDashboard([rep(0, 'skipped'), rep(1, 'skipped'), rep(2, 'skipped'), rep(3, 'skipped')])
    await waitFor(() => m.store().reports.length === 0, 'all rows hidden')
    assert.equal(m.store().activeId, '')
    assert.equal(m.store().activeReport, null)
    assert.deepEqual(m.store().selectedIds, [])
    assert.deepEqual(mutationCalls(m.state.calls), [], 'active-row reconciliation issues no mutation or delete bridge call')
  } finally {
    await m.cleanup()
  }
})

test('a skipped report that later becomes actionable again reappears: no permanent path blacklist', async () => {
  const m = await mount([rep(0, 'skipped'), rep(1, 'waiting')])
  try {
    sameList(m.ids(), ['1'])
    m.setDashboard([rep(0, 'waiting'), rep(1, 'waiting')])
    await waitFor(() => m.ids().join(',') === '0,1', 'row reappears after a legitimate status change')
    assert.equal(m.store().stats.total, 2)
    assert.deepEqual(mutationCalls(m.state.calls), [])
  } finally {
    await m.cleanup()
  }
})

test('select-all in the rendered table never includes skipped reports', async () => {
  const m = await mount([rep(0, 'waiting'), rep(1, 'skipped'), rep(2, 'completed')])
  try {
    const selectAll = () => m.document().querySelector('[aria-label="Chọn tất cả"]')
    assert.ok(selectAll(), 'select-all control rendered')
    await m.act(() => selectAll().dispatchEvent(new m.MouseEvent('click', { bubbles: true })))
    sameList(m.store().selectedIds, ['0', '2'])
    assert.equal(m.store().selectedIds.includes('1'), false)
    await m.act(() => selectAll().dispatchEvent(new m.MouseEvent('click', { bubbles: true })))
    sameList(m.store().selectedIds, [])
  } finally {
    await m.cleanup()
  }
})

test('visible totals stay consistent with the rendered rows after hiding skipped reports', async () => {
  const m = await mount([rep(0, 'waiting'), rep(1, 'skipped'), rep(2, 'needs_review'), rep(3, 'error'), rep(4, 'completed')])
  try {
    const s = m.store()
    assert.equal(s.stats.total, s.reports.length)
    assert.equal(s.stats.total, m.rows().length)
    assert.equal(s.stats.completed + s.stats.needs_review + s.stats.error + s.stats.skipped, 3)
    assert.match(m.text(), /4\/4 file/)
    assert.match(m.text(), /Tổng số báo cáo: 4/)
  } finally {
    await m.cleanup()
  }
})

test('1000+ mixed reports render the exact visible count without pathological reconciliation', async () => {
  const total = 1_200
  const reports = Array.from({ length: total }, (_, i) => rep(i, i % 4 === 0 ? 'skipped'
    : ['waiting', 'completed', 'needs_review', 'error'][i % 3]))
  const expected = total - total / 4
  const m = await mount(reports)
  try {
    await waitFor(() => m.store().reports.length === expected, 'bulk visible count', 8_000)
    assert.equal(m.store().stats.total, expected)
    assert.ok(m.store().reports.every((report) => report.status !== 'skipped'))
    assert.equal(m.rows().length, expected)
    assert.match(m.text(), new RegExp(`${expected}/${expected} file`))
    // Reconciliation is one dashboard read per poll: no per-report bridge calls for hidden rows.
    assert.equal(m.state.calls.filter((name) => name === 'get_report_details').length <= 1, true)
    assert.deepEqual(mutationCalls(m.state.calls), [])
  } finally {
    await m.cleanup()
  }
})
