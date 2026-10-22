/**
 * PROMPT-040 — real-browser harness for the remembered Excel files (“Excel mặc định”).
 *
 * The React/state layer is the real application (`App` + `StoreProvider`, real `SourcePanel`); only the
 * pywebview bridge is mocked. It exists because the claims under test are about rendered UI behaviour:
 *   • a restored default is visible (name + folder) without any picker call on startup,
 *   • a stored-but-missing/invalid file is reported instead of silently replaced,
 *   • “Đổi file…” opens exactly one native picker and shows the newly persisted path,
 *   • “Xóa lựa chọn” calls the backend and returns the row to “Chưa chọn”.
 * `window.__tnpExcelHarness` lets the driver change the backend state between assertions.
 */
import React from 'react'
import ReactDOM from 'react-dom/client'
import App from '../../src/App'
import { StoreProvider } from '../../src/state/store'
import '../../src/index.css'
import type { ExcelKindState } from '../../src/types'

type NativeEnvelope = { ok: true; data: unknown }

const TEMPLATE_LABEL = 'File kiểm chứng Excel'
const OUTPUT_LABEL = 'File kết quả Excel'

const templateDefault: ExcelKindState = {
  kind: 'template',
  label: TEMPLATE_LABEL,
  path: 'D:\\Kiểm chứng 2026\\Biểu mẫu kiểm chứng.xlsx',
  name: 'Biểu mẫu kiểm chứng.xlsx',
  folder: 'D:\\Kiểm chứng 2026',
  state: 'ok',
  message: 'Hợp lệ: OK – sheet \'Kiểm chứng\' (12 cột nhận diện)',
  persisted: true,
}
const outputDefault: ExcelKindState = {
  kind: 'output',
  label: OUTPUT_LABEL,
  path: 'D:\\Kết quả\\Kiem chung 09-2026.xlsx',
  name: 'Kiem chung 09-2026.xlsx',
  folder: 'D:\\Kết quả',
  state: 'ok',
  message: 'Đang dùng File kết quả đã lưu.',
  persisted: true,
}

let excel: { template: ExcelKindState; output: ExcelKindState } = {
  template: { ...templateDefault },
  output: { ...outputDefault },
}
/** What the next native pick of each kind returns (null = the user cancels the dialog). */
const pendingPicks: Record<string, string | null> = {}
const calls: Record<string, number> = {}
const cleared: string[] = []

const unconfigured = (kind: 'template' | 'output'): ExcelKindState => ({
  kind,
  label: kind === 'template' ? TEMPLATE_LABEL : OUTPUT_LABEL,
  path: '',
  name: '',
  folder: '',
  state: 'unconfigured',
  message: kind === 'template' ? 'Chưa chọn File kiểm chứng Excel.' : 'Chưa chọn File kết quả Excel.',
  persisted: true,
})

const envelope = (data: unknown): NativeEnvelope => ({ ok: true, data })

const appDto = () => ({ name: 'Report Extractor', title: 'Report Extractor', version: '1.4.0', build: '021',
  buildNumber: 21, runtime: 'pywebview-development', updateInstallCapable: false })

const configDto = () => ({
  paths: { reportFolder: 'D:\\Reports', template: excel.template.path, output: excel.output.path },
  period: { mode: 'auto', month: '', year: '', from: '', to: '' },
  forceReprocess: false,
  ollama: { host: '', port: '11434', model: '', checked: 'unchecked', message: '', models: [] },
  update: { path: '', autoCheck: false },
})

const dashboardDto = () => ({
  app: appDto(),
  config: configDto(),
  excel,
  scan: { scanned: true, message: 'harness', generation: 1, revision: 1, inProgress: false,
    incremental: false, completedLeaves: 0, totalLeaves: 0, firstLeafMs: null },
  reports: [],
  job: { status: 'idle', queue: [], index: 0, doneCount: 0, stage: 'waiting', percent: 0,
    currentFile: '', elapsedSec: 0, remainSec: null, startedAt: null, finishedAt: null, hasSamples: false },
  logs: [],
  ollama: { host: '', port: '11434', model: '', checked: 'unchecked', message: '', models: [], aiStatus: '',
    discovering: false, discoveryMessage: '', discoveryResults: [], discoveryChecked: 0, discoveryTotal: 0,
    serverApplyRunning: false, serverApplyMessage: '' },
  update: { path: '', autoCheck: false, status: 'idle', message: '', progress: 0, progressStage: '',
    installEnabled: false, canInstall: false, installDisabledReason: '', error: '', available: null },
  diagnostics: { running: false, rows: [], error: '' },
})

function pick(kind: 'template' | 'output'): unknown {
  const chosen = pendingPicks[kind] ?? null
  if (chosen === null) return { cancelled: true }
  const name = chosen.split(/[\\/]/).pop() ?? chosen
  const folder = chosen.slice(0, Math.max(0, chosen.length - name.length - 1))
  excel = { ...excel, [kind]: { ...excel[kind], path: chosen, name, folder, state: 'ok', persisted: true,
    message: kind === 'template' ? 'Hợp lệ: OK – sheet \'Kiểm chứng\' (12 cột nhận diện)' : 'Đang dùng File kết quả đã lưu.' } }
  return kind === 'template' ? { cancelled: false, path: chosen, kind: 'template' } : { cancelled: false, path: chosen }
}

const handlers: Record<string, (...args: unknown[]) => unknown> = {
  ping: () => 'ok',
  get_app_version: () => appDto(),
  get_current_config: () => configDto(),
  get_dashboard_state: () => dashboardDto(),
  get_excel_selection: () => excel,
  choose_template_file: () => pick('template'),
  choose_output_file: () => pick('output'),
  clear_excel_selection: (kind: unknown) => {
    const key = kind === 'output' ? 'output' : 'template'
    cleared.push(key)
    excel = { ...excel, [key]: unconfigured(key) }
    return excel
  },
}

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
    __tnpExcelHarness?: {
      bridgeCalls: () => Record<string, number>
      clearedKinds: () => string[]
      excel: () => typeof excel
      setExcel: (next: { template?: Partial<ExcelKindState>; output?: Partial<ExcelKindState> }) => void
      queuePick: (kind: 'template' | 'output', path: string | null) => void
    }
  }
}

window.pywebview = { api }
window.__tnpExcelHarness = {
  bridgeCalls: () => ({ ...calls }),
  clearedKinds: () => [...cleared],
  excel: () => excel,
  setExcel(next) {
    excel = {
      template: { ...excel.template, ...(next.template ?? {}) },
      output: { ...excel.output, ...(next.output ?? {}) },
    }
  },
  queuePick(kind, path) {
    pendingPicks[kind] = path
  },
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
