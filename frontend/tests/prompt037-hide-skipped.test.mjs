// PROMPT-037 / PROMPT-037R regression: every status whose UI meaning is “Bỏ qua” — skipped,
// outside_period, source_duplicate, fast_skip — is hidden from “Danh sách báo cáo” and from its
// counters, selection and active row. The DTO keeps the rows; the presentation rule is the only
// filter. Store-level tests drive the real store through a mocked pywebview bridge; the rendered
// ReportTable and StatsCard are checked against the store.
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
let FILTERS
let presentReports
let isVisibleReport
let HIDDEN_REPORT_STATUSES
let BUCKET_OF
let STATUS_LABEL
let ReportTable
let StatsCard

const BO_QUA_STATUSES = ['skipped', 'outside_period', 'source_duplicate', 'fast_skip']
const VISIBLE_STATUSES = ['waiting', 'new_row', 'processing', 'completed', 'needs_review', 'error', 'cancelled', 'excluded']

before(async () => {
  server = await createServer({
    configFile: `${frontendRoot}/vite.config.ts`,
    root: frontendRoot,
    appType: 'custom',
    server: { middlewareMode: true, hmr: false },
  })
  ;({ StoreProvider, useStore, resolveActiveReportId, FILTERS } = await server.ssrLoadModule('/src/state/store.tsx'))
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

// Deterministic mixed population: every 4th row is a “Bỏ qua” status (cycling the four), others visible.
function mixed(total) {
  return Array.from({ length: total }, (_, i) => rep(i, i % 4 === 0
    ? BO_QUA_STATUSES[(i / 4) % BO_QUA_STATUSES.length]
    : VISIBLE_STATUSES.slice(0, 4)[i % 4]))
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
    setDashboard,
    document: () => dom.window.document,
    MouseEvent: dom.window.MouseEvent,
    rows: () => [...dom.window.document.querySelectorAll('tbody tr')],
    text: () => dom.window.document.body.textContent,
    act: (fn) => React.act(async () => { await fn() }),
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
  for (const status of VISIBLE_STATUSES) {
    assert.equal(isVisibleReport({ status }), true, `${status} must stay visible`)
  }
})

test('visibility rule: every “Bỏ qua” status is hidden: skipped, outside_period, source_duplicate, fast_skip', () => {
  for (const status of BO_QUA_STATUSES) {
    assert.equal(isVisibleReport({ status }), false, `${status} must be hidden`)
  }
  sameList([...HIDDEN_REPORT_STATUSES].sort(), [...BO_QUA_STATUSES].sort())
})

test('the “File bị bỏ qua” filter is removed because no visible row could ever match it', () => {
  assert.equal(FILTERS.includes('File bị bỏ qua'), false)
  assert.equal(FILTERS.length, 6)
})

test('mixed list: only “Bỏ qua” reports are removed and STT is renumbered over the visible rows', () => {
  const input = [rep(0, 'waiting'), rep(1, 'skipped'), rep(2, 'completed'), rep(3, 'outside_period'),
    rep(4, 'needs_review'), rep(5, 'source_duplicate'), rep(6, 'error'), rep(7, 'fast_skip')]
  const out = presentReports(input)
  sameList(out.map((r) => r.id), ['0', '2', '4', '6'])
  sameList(out.map((r) => r.stt), [1, 2, 3, 4])
  // Unchanged rows keep their identity; only STT-changed rows are copied.
  assert.equal(out[0], input[0])
  assert.notEqual(out[1], input[2])
  assert.equal(out[1].stt, 2)
})

test('1000+ mixed reports: exact visible count and linear, non-pathological cost', () => {
  const total = 12_000
  const input = mixed(total)
  const started = performance.now()
  const out = presentReports(input)
  const elapsed = performance.now() - started
  assert.equal(out.length, total - total / 4)
  assert.ok(out.every((r) => isVisibleReport(r)))
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
  assert.equal(BUCKET_OF.outside_period, 'skipped')
  assert.equal(BUCKET_OF.source_duplicate, 'skipped')
  assert.equal(BUCKET_OF.fast_skip, 'skipped')
  assert.equal(STATUS_LABEL.skipped, 'Bỏ qua — đã cập nhật')
  assert.equal(STATUS_LABEL.outside_period, 'Bỏ qua ngoài thời gian xử lý')
  assert.equal(STATUS_LABEL.fast_skip, 'Bỏ qua nhanh — đã xử lý gần đây')
})

// ---------------------------------------------------------------- store + rendered list

test('a full dashboard hides “Bỏ qua” rows; waiting/completed/needs_review/error stay; header and totals match the list', async () => {
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

test('every “Bỏ qua” status is hidden from rows, totals, selection, active row and select-all together', async () => {
  const m = await mount([rep(0, 'waiting'), rep(1, 'skipped'), rep(2, 'outside_period'), rep(3, 'source_duplicate'),
    rep(4, 'fast_skip'), rep(5, 'completed')])
  try {
    sameList(m.ids(), ['0', '5'])
    assert.equal(m.store().stats.total, 2)
    assert.equal(m.rows().length, 2)
    assert.match(m.text(), /2\/2 file/)
    assert.doesNotMatch(m.text(), /report-[1-4]\.pptx/)
    const selectAll = () => m.document().querySelector('[aria-label="Chọn tất cả"]')
    await m.act(() => selectAll().dispatchEvent(new m.MouseEvent('click', { bubbles: true })))
    sameList(m.store().selectedIds, ['0', '5'])
    await m.act(() => selectAll().dispatchEvent(new m.MouseEvent('click', { bubbles: true })))
    sameList(m.store().selectedIds, [])
    // Hidden ids cannot be selected or made active through the store: the store ignores them on the next reconciliation.
    await m.act(() => { m.store().toggleSelected('2'); m.store().setActiveId('3') })
    m.setDashboard([rep(0, 'waiting'), rep(1, 'skipped'), rep(2, 'outside_period'), rep(3, 'source_duplicate'),
      rep(4, 'fast_skip'), rep(5, 'completed')])
    await waitFor(() => m.store().selectedIds.length === 0 && m.store().activeId === '0', 'hidden selection and active pruned')
    assert.deepEqual(mutationCalls(m.state.calls), [], 'no mutation or delete bridge call')
  } finally {
    await m.cleanup()
  }
})

test('a visible row turning into any “Bỏ qua” status leaves the list on the next poll and stays out', async () => {
  const m = await mount([rep(0, 'waiting'), rep(1, 'waiting'), rep(2, 'waiting'), rep(3, 'waiting')])
  try {
    m.setDashboard([rep(0, 'outside_period'), rep(1, 'waiting'), rep(2, 'source_duplicate'), rep(3, 'fast_skip')])
    await waitFor(() => m.ids().join(',') === '1', 'three skip kinds removed by poll')
    const reads = m.state.dashboardReads
    await m.act(() => delay(1_600))
    assert.ok(m.state.dashboardReads >= reads + 2, 'dashboard polling continued')
    sameList(m.ids(), ['1'])
    assert.deepEqual(mutationCalls(m.state.calls), [], 'no mutation or delete bridge call')
  } finally {
    await m.cleanup()
  }
})

test('incremental deltas never reinsert outside_period, source_duplicate or fast_skip; final scan keeps them hidden', async () => {
  const m = await mount([rep(0, 'waiting'), rep(1, 'waiting')])
  try {
    let resolveScan
    m.state.scanReply = new Promise((resolve) => { resolveScan = resolve })
    await m.act(() => { void m.store().scan() })
    await waitFor(() => m.store().scanBusy === true, 'scan started')
    m.state.dashboard = dashboard({ scanned: true, message: 'leaf 1', generation: 2, revision: 1, inProgress: true,
      incremental: true, completedLeaves: 1, totalLeaves: 2 },
    [rep(0, 'fast_skip'), rep(2, 'outside_period'), rep(3, 'waiting'), rep(4, 'source_duplicate')])
    // scan() clears the list; the delta alone repopulates it: skip kinds never enter, visible rows do.
    await waitFor(() => m.ids().join(',') === '3', 'delta keeps only visible rows')
    await m.act(() => delay(900))
    sameList(m.ids(), ['3'])
    resolveScan(dashboard({ scanned: true, message: 'done', generation: 2, revision: 2, inProgress: false },
      [rep(0, 'fast_skip'), rep(1, 'waiting'), rep(2, 'outside_period'), rep(3, 'waiting'), rep(4, 'source_duplicate')]))
    await waitFor(() => m.store().scanBusy === false, 'final scan response')
    sameList(m.ids(), ['1', '3'])
    assert.equal(m.store().stats.total, 2)
    assert.deepEqual(mutationCalls(m.state.calls), [], 'no mutation or delete bridge call')
  } finally {
    await m.cleanup()
  }
})

test('incremental deltas remove a visible row that transitions to any “Bỏ qua” status inside a delta', async () => {
  const m = await mount([rep(0, 'waiting'), rep(1, 'waiting'), rep(2, 'waiting'), rep(3, 'waiting')])
  try {
    let resolveScan
    m.state.scanReply = new Promise((resolve) => { resolveScan = resolve })
    await m.act(() => { void m.store().scan() })
    await waitFor(() => m.store().scanBusy === true, 'scan started')
    m.state.dashboard = dashboard({ scanned: true, message: 'leaf', generation: 2, revision: 1, inProgress: true,
      incremental: true, completedLeaves: 1, totalLeaves: 2 },
    [rep(0, 'skipped'), rep(1, 'outside_period'), rep(2, 'source_duplicate'), rep(3, 'fast_skip')])
    await waitFor(() => m.ids().length === 0, 'every transitioned row removed')
    resolveScan(dashboard({ scanned: true, message: 'done', generation: 2, revision: 2, inProgress: false },
      [rep(0, 'skipped'), rep(1, 'outside_period'), rep(2, 'source_duplicate'), rep(3, 'fast_skip')]))
    await waitFor(() => m.store().scanBusy === false, 'final scan response')
    assert.deepEqual(m.ids(), [])
    assert.equal(m.store().stats.total, 0)
    assert.equal(m.store().activeId, '')
  } finally {
    await m.cleanup()
  }
})

test('a rescan with the same “Bỏ qua” classification stays hidden', async () => {
  const m = await mount([rep(0, 'waiting'), rep(1, 'skipped'), rep(2, 'completed'), rep(3, 'fast_skip')])
  try {
    sameList(m.ids(), ['0', '2'])
    m.state.scanReply = dashboard({ scanned: true, message: 'rescan', generation: 2, revision: 1, inProgress: false },
      [rep(0, 'waiting'), rep(1, 'skipped'), rep(2, 'completed'), rep(3, 'fast_skip')])
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

test('a selected report that becomes “Bỏ qua” is removed from selectedIds', async () => {
  const m = await mount([rep(0, 'waiting'), rep(1, 'waiting'), rep(2, 'completed')])
  try {
    await m.act(() => { m.store().toggleSelected('0'); m.store().toggleSelected('2') })
    sameList(m.store().selectedIds, ['0', '2'])
    m.setDashboard([rep(0, 'waiting'), rep(1, 'waiting'), rep(2, 'fast_skip')])
    await waitFor(() => m.store().selectedIds.length === 1, 'hidden selection pruned')
    sameList(m.store().selectedIds, ['0'])
  } finally {
    await m.cleanup()
  }
})

test('an active row that becomes “Bỏ qua” advances to the next visible row; the last one falls back to the previous; empty clears', async () => {
  const m = await mount([rep(0, 'waiting'), rep(1, 'waiting'), rep(2, 'waiting'), rep(3, 'waiting')])
  try {
    await m.act(() => m.store().setActiveId('1'))
    assert.equal(m.store().activeId, '1')
    m.setDashboard([rep(0, 'waiting'), rep(1, 'outside_period'), rep(2, 'waiting'), rep(3, 'waiting')])
    await waitFor(() => m.store().activeId === '2', 'active advanced to next visible row')
    assert.equal(m.store().activeReport?.id, '2')

    await m.act(() => m.store().setActiveId('3'))
    m.setDashboard([rep(0, 'waiting'), rep(1, 'outside_period'), rep(2, 'waiting'), rep(3, 'source_duplicate')])
    await waitFor(() => m.store().activeId === '2', 'active fell back to previous visible row')

    m.setDashboard([rep(0, 'skipped'), rep(1, 'outside_period'), rep(2, 'source_duplicate'), rep(3, 'fast_skip')])
    await waitFor(() => m.store().reports.length === 0, 'all rows hidden')
    assert.equal(m.store().activeId, '')
    assert.equal(m.store().activeReport, null)
    assert.deepEqual(m.store().selectedIds, [])
    assert.deepEqual(mutationCalls(m.state.calls), [], 'active-row reconciliation issues no mutation or delete bridge call')
  } finally {
    await m.cleanup()
  }
})

test('a “Bỏ qua” report that later becomes actionable again reappears: no permanent path blacklist', async () => {
  const m = await mount([rep(0, 'skipped'), rep(1, 'outside_period'), rep(2, 'fast_skip'), rep(3, 'waiting')])
  try {
    sameList(m.ids(), ['3'])
    m.setDashboard([rep(0, 'waiting'), rep(1, 'waiting'), rep(2, 'waiting'), rep(3, 'waiting')])
    await waitFor(() => m.ids().join(',') === '0,1,2,3', 'rows reappear after a legitimate status change')
    assert.equal(m.store().stats.total, 4)
    assert.deepEqual(mutationCalls(m.state.calls), [])
  } finally {
    await m.cleanup()
  }
})

test('select-all in the rendered table never includes “Bỏ qua” reports', async () => {
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

test('visible totals stay consistent with the rendered rows after hiding “Bỏ qua” reports', async () => {
  const m = await mount([rep(0, 'waiting'), rep(1, 'skipped'), rep(2, 'needs_review'), rep(3, 'error'),
    rep(4, 'completed'), rep(5, 'outside_period')])
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
  const reports = mixed(total)
  const expected = total - total / 4
  const m = await mount(reports)
  try {
    await waitFor(() => m.store().reports.length === expected, 'bulk visible count', 8_000)
    assert.equal(m.store().stats.total, expected)
    assert.ok(m.store().reports.every((report) => isVisibleReport(report)))
    assert.equal(m.rows().length, expected)
    assert.match(m.text(), new RegExp(`${expected}/${expected} file`))
    // Reconciliation is one dashboard read per poll: no per-report bridge calls for hidden rows.
    assert.ok(m.state.calls.filter((name) => name === 'get_report_details').length <= 1)
    assert.deepEqual(mutationCalls(m.state.calls), [])
  } finally {
    await m.cleanup()
  }
})
