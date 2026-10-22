/**
 * PROMPT-034R — real-browser layout harness for the “Danh sách báo cáo” viewport.
 *
 * The React/state layer is the real application (`App` + `StoreProvider`), only the pywebview bridge
 * is mocked. It exists because jsdom performs no layout: proving “~6 visible rows”, “the column header
 * stays sticky while rows scroll”, “1000 reports do not grow the page” and “no horizontal scrollbar at
 * desktop width” requires measured geometry from a real engine.
 *
 * Usage (see `tests/prompt034r-report-list-layout.test.mjs`):
 *   /tests/layout/report-list-harness.html?rows=1000
 *
 * `window.__tnpLayoutHarness` lets the driver model scan updates on a running page, so one session can
 * cover the ≤6-row case, the >6-row case and an incremental scan leaf without reloads.
 */
import React from 'react'
import ReactDOM from 'react-dom/client'
import App from '../../src/App'
import { StoreProvider } from '../../src/state/store'
import '../../src/index.css'

type NativeEnvelope = { ok: true; data: unknown }
type ReportRow = Record<string, unknown>

const params = new URLSearchParams(window.location.search)
const initialRows = Math.max(0, Math.min(5000, Number(params.get('rows') ?? '6') || 0))

const STATUS_CYCLE = ['completed', 'completed', 'needs_review', 'waiting', 'error'] as const

function report(index: number): ReportRow {
  const status = STATUS_CYCLE[index % STATUS_CYCLE.length]
  return {
    id: `r${index}`,
    stt: index + 1,
    managementNumber: `2609${String(index + 1).padStart(5, '0')}`,
    fileName: `(CTMS)_2609${String(index + 1).padStart(5, '0')}_harness.pptx`,
    path: '',
    occurrenceDate: '01/10/2026',
    vendor: index % 3 === 0 ? 'Vendor A' : '',
    manualFields: {},
    slides: 12,
    status,
    shortResult: 'Nguyên nhân 2 · Cải tiến 3 · Ảnh 4',
    results: { causes: [], countermeasures: [], temporaryRemoved: false, afterImages: [] },
  }
}

function dataset(count: number): ReportRow[] {
  return Array.from({ length: count }, (_, index) => report(index))
}

let rows = dataset(initialRows)
let revision = 1
/** Armed by an incremental harness action; consumed by the next dashboard snapshot, as Python does. */
let incrementalRows: ReportRow[] | null = null

const appDto = () => ({
  name: 'Report Extractor',
  title: 'Report Extractor',
  version: '1.4.0',
  build: '021',
  buildNumber: 21,
})

const configDto = () => ({
  paths: { reportFolder: 'D:\\Reports', template: '', output: 'D:\\Reports\\Output\\ket_qua.xlsx' },
  period: { mode: 'auto', month: '', year: '', from: '', to: '' },
  forceReprocess: false,
  ollama: { host: '', port: '11434', model: '', checked: 'unchecked', message: '', models: [] },
  update: { path: '', autoCheck: false },
})

function dashboardDto() {
  const incremental = incrementalRows !== null
  const snapshot = incrementalRows ?? rows
  incrementalRows = null
  return {
    app: appDto(),
    config: configDto(),
    scan: {
      scanned: true,
      message: `harness · ${rows.length} báo cáo`,
      generation: 1,
      revision,
      inProgress: false,
      incremental,
      completedLeaves: incremental ? rows.length : 0,
      totalLeaves: rows.length,
      firstLeafMs: null,
    },
    reports: snapshot,
    job: { status: 'idle', queue: [], index: 0, doneCount: 0, stage: 'waiting', percent: 0,
      currentFile: '', elapsedSec: 0, remainSec: null, startedAt: null, finishedAt: null, hasSamples: false },
    logs: [],
    ollama: { host: '', port: '11434', model: '', checked: 'unchecked', message: '', models: [], aiStatus: '',
      discovering: false, discoveryMessage: '', discoveryResults: [], discoveryChecked: 0, discoveryTotal: 0,
      serverApplyRunning: false, serverApplyMessage: '' },
    update: { path: '', autoCheck: false, status: 'idle', message: '', progress: 0, progressStage: '',
      installEnabled: false, canInstall: false, installDisabledReason: '', error: '', available: null },
    diagnostics: { running: false, rows: [], error: '' },
  }
}

const envelope = (data: unknown): NativeEnvelope => ({ ok: true, data })

/** Every method the running store can reach is answered; unknown ones resolve to `null`. */
const handlers: Record<string, (...args: unknown[]) => unknown> = {
  ping: () => 'ok',
  get_app_version: () => appDto(),
  get_current_config: () => configDto(),
  get_dashboard_state: () => dashboardDto(),
  get_report_details: (id: unknown) => ({
    id,
    causes: ['Nguyên nhân giữ nguyên văn từ report nguồn.'],
    countermeasures: ['Đối sách cải tiến giữ nguyên văn.'],
    afterImages: [],
    qpn: undefined,
    slides: 12,
    excelRow: 128,
    temporaryRemoved: false,
    warning: '',
    occurrenceDate: '01/10/2026',
    vendor: '',
  }),
}

/** Native traffic tally — PROMPT-034R must not add per-row bridge calls. */
const calls: Record<string, number> = {}

const api = new Proxy(handlers, {
  get(target, property: string) {
    const handler = target[property]
    return (...args: unknown[]) => {
      calls[property] = (calls[property] ?? 0) + 1
      return Promise.resolve(envelope(handler ? handler(...args) : null))
    }
  },
})

declare global {
  interface Window {
    pywebview?: { api: unknown }
    __tnpLayoutHarness?: {
      appendIncrementalRows: (count: number) => number
      setStatus: (index: number, status: string) => number
      rowCount: () => number
      bridgeCalls: () => Record<string, number>
    }
  }
}

window.pywebview = { api }
window.__tnpLayoutHarness = {
  /** Emulates one completed incremental scan leaf: only the newly discovered reports travel in the snapshot. */
  appendIncrementalRows(count) {
    const start = rows.length
    const added = Array.from({ length: count }, (_, offset) => report(start + offset))
    rows = rows.concat(added)
    revision += 1
    incrementalRows = added
    return rows.length
  },
  setStatus(index, status) {
    if (rows[index]) rows[index] = { ...rows[index], status }
    revision += 1
    return revision
  },
  rowCount: () => rows.length,
  bridgeCalls: () => ({ ...calls }),
}

const rootElement = document.getElementById('root')
if (!rootElement) throw new Error('harness root element is missing')

ReactDOM.createRoot(rootElement).render(
  <React.StrictMode>
    <StoreProvider>
      <App />
    </StoreProvider>
  </React.StrictMode>,
)
