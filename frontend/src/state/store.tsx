import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from 'react'
import { BridgeCallError, callBridge, hasNativeBridge } from '../lib/bridge'
import {
  BUCKET_OF,
  type ContentCandidate,
  type DiagnosticRow,
  type DiscoveredServer,
  type ImageCandidate,
  type LogEntry,
  type OllamaConnection,
  type ProcessingRun,
  type Report,
  type StageKey,
  type TabKey,
} from '../types'

export const FILTERS = [
  'Tất cả file đã quét',
  'File cần xử lý',
  'File đã xử lý',
  'File cần kiểm tra',
  'File lỗi',
  'File bị bỏ qua',
  'File đã loại thủ công',
] as const
export type FilterKey = (typeof FILTERS)[number]
export type PeriodMode = 'auto' | 'all' | 'month' | 'range'

export type UpdateState = {
  path: string
  autoCheck: boolean
  status: 'idle' | 'checking' | 'available' | 'latest' | 'error' | 'installing' | 'installed'
  message: string
  available?: { version: string; build: string; package: string; size: string }
  progress: number
  progressStage: string
  installEnabled?: boolean
}

export type LearningState = {
  images: ImageCandidate[]
  contents: ContentCandidate[]
  counts: { image: { total: number; labeled: number }; content: { total: number; labeled: number } }
  modelStatus: { image: string; content: string }
  excelPending: number
  excelLastResult: { kind: 'ok' | 'locked' | 'none'; message: string }
  training?: boolean
  trainingMessage?: string
  reviewSummary?: string
  contentReviewSummary?: string
}

type NativeConfig = {
  paths: { reportFolder: string; template: string; output: string }
  period: { mode: PeriodMode; month: string; year: string; from: string; to: string }
  forceReprocess: boolean
  ollama: { host: string; port: string; model: string; checked?: OllamaConnection['checked'];
    message?: string; models?: string[] }
  update: { path: string; autoCheck: boolean }
  app?: { name: string; title: string; version: string; build: string; buildNumber: number }
}

type DashboardDTO = {
  app: { name: string; title: string; version: string; build: string; buildNumber: number }
  config: NativeConfig
  scan: { scanned: boolean; message: string }
  reports: Report[]
  job: ProcessingRun & { error?: string }
  logs: LogEntry[]
  ollama: {
    host: string
    port: string
    model: string
    checked: OllamaConnection['checked']
    message: string
    models: string[]
    aiStatus: string
    discovering: boolean
    discoveryMessage: string
    discoveryResults: DiscoveredServer[]
    discoveryChecked: number
    discoveryTotal: number
    serverApplyRunning: boolean
    serverApplyMessage: string
  }
  update: UpdateState
  diagnostics: { running: boolean; rows: DiagnosticRow[]; error: string }
}

const EMPTY_RUN: ProcessingRun = {
  status: 'idle', queue: [], index: 0, doneCount: 0, stage: 'waiting', percent: 0, currentFile: '',
  elapsedSec: 0, remainSec: null, startedAt: null, finishedAt: null, hasSamples: false,
}
const EMPTY_LEARNING: LearningState = {
  images: [], contents: [],
  counts: { image: { total: 0, labeled: 0 }, content: { total: 0, labeled: 0 } },
  modelStatus: { image: 'Chưa tải trạng thái mô hình.', content: 'Chưa tải trạng thái mô hình.' },
  excelPending: 0,
  excelLastResult: { kind: 'none', message: '' },
}
const EMPTY_OLLAMA: OllamaConnection = {
  server: '', port: '11434', model: '', checked: 'unchecked', models: [], message: 'Chưa kiểm tra kết nối.',
}
export const REVEAL_DONE = 5

type StoreValue = ReturnType<typeof useStoreValue>
const StoreContext = createContext<StoreValue | null>(null)

function stageKey(value: unknown): StageKey {
  const valid: StageKey[] = ['waiting', 'reading', 'analyzing', 'analyzing_heuristic', 'extracting',
    'extracting_qpn', 'extracting_images', 'writing_excel']
  return valid.includes(value as StageKey) ? value as StageKey : 'waiting'
}

function reportFromDto(value: Report): Report {
  return {
    ...value,
    id: String(value.id),
    path: '', // absolute source paths remain inside Python; fileName is the display identity
    warning: value.warning || undefined,
    vendor: value.vendor || '',
    manualFields: value.manualFields || {},
    results: value.results || { causes: [], countermeasures: [], temporaryRemoved: false, afterImages: [] },
  }
}

function normalizeDashboard(raw: DashboardDTO): DashboardDTO {
  const reports = (raw.reports || []).map(reportFromDto)
  return {
    ...raw,
    reports,
    job: { ...EMPTY_RUN, ...(raw.job || {}), stage: stageKey(raw.job?.stage) },
    ollama: { ...EMPTY_OLLAMA, ...(raw.ollama || {}) },
    diagnostics: raw.diagnostics || { running: false, rows: [], error: '' },
    logs: raw.logs || [],
  }
}

function useStoreValue() {
  const [tabValue, setTabValue] = useState<TabKey>('reports')
  const [connected, setConnected] = useState(false)
  const [errorMessage, setErrorMessage] = useState('')
  const [errorCode, setErrorCode] = useState('')
  const [appInfo, setAppInfo] = useState({ name: 'Report Extractor', title: 'Report Extractor', version: '', build: '' })
  const [reports, setReports] = useState<Report[]>([])
  const [scanned, setScanned] = useState(false)
  const [scanDirty, setScanDirty] = useState(false)
  const [scanBusy, setScanBusy] = useState(false)
  const [scanMessage, setScanMessage] = useState('Chưa chọn thư mục báo cáo PPTX.')
  const [selectedIds, setSelectedIds] = useState<string[]>([])
  const [activeId, setActiveIdValue] = useState('')
  const [search, setSearch] = useState('')
  const [filter, setFilter] = useState<FilterKey>('Tất cả file đã quét')
  const [folder, setFolderValue] = useState('')
  const [template, setTemplateValue] = useState('')
  const [output, setOutputValue] = useState('')
  const [periodMode, setPeriodModeValue] = useState<PeriodMode>('auto')
  const [month, setMonthValue] = useState('')
  const [year, setYearValue] = useState('')
  const [fromDate, setFromDateValue] = useState('')
  const [toDate, setToDateValue] = useState('')
  const [force, setForceValue] = useState(false)
  const [run, setRun] = useState<ProcessingRun>(EMPTY_RUN)
  const [reveal, setReveal] = useState<Record<string, number>>({})
  const [log, setLog] = useState<LogEntry[]>([])
  const [logCollapsed, setLogCollapsed] = useState(false)
  const clearLogThroughRef = useRef(0)
  const [ollama, setOllamaState] = useState<OllamaConnection>(EMPTY_OLLAMA)
  const [aiStatusText, setAiStatusText] = useState('')
  const [discovering, setDiscovering] = useState(false)
  const [discoveryMessage, setDiscoveryMessage] = useState('LAN discovery only runs when requested.')
  const [discoveryResults, setDiscoveryResults] = useState<DiscoveredServer[]>([])
  const [appliedServer, setAppliedServer] = useState<string | null>(null)
  const [serverApplyRunning, setServerApplyRunning] = useState(false)
  const [serverApplyMessage, setServerApplyMessage] = useState('')
  const [update, setUpdateState] = useState<UpdateState>({
    path: '', autoCheck: true, status: 'idle', message: 'Chưa kiểm tra cập nhật.', progress: 0, progressStage: '',
  })
  const [diagRows, setDiagRows] = useState<DiagnosticRow[]>([])
  const [diagRunning, setDiagRunning] = useState(false)
  const [learningState, setLearningState] = useState<LearningState>(EMPTY_LEARNING)
  const [imageReviewIndex, setImageReviewIndex] = useState(0)
  const [contentReviewIndex, setContentReviewIndex] = useState(0)
  const [training, setTraining] = useState(false)
  const [inputFileToken, setInputFileToken] = useState('')
  const [inputFileName, setInputFileName] = useState('')
  const dashboardInFlight = useRef(false)
  const learningInFlight = useRef(false)
  const lastLearningRefresh = useRef(0)
  const tabRef = useRef<TabKey>(tabValue)
  const detailStatus = useRef<Record<string, string>>({})
  const ollamaDirty = useRef(false)
  const updateDirty = useRef(false)
  const alive = useRef(true)
  const recordError = useCallback((error: unknown) => {
    const message = error instanceof Error ? error.message : String(error)
    const code = error instanceof BridgeCallError ? error.code
      : message.includes('Python bridge is not connected') ? 'BRIDGE_UNAVAILABLE' : 'INTERNAL_ERROR'
    setErrorCode(code)
    setErrorMessage(message)
  }, [])
  const clearError = useCallback(() => { setErrorCode(''); setErrorMessage('') }, [])

  const configDto = useCallback((): NativeConfig => ({
    paths: { reportFolder: folder, template, output },
    period: { mode: periodMode, month, year, from: fromDate, to: toDate },
    forceReprocess: force,
    ollama: { host: ollama.server, port: ollama.port, model: ollama.model },
    update: { path: update.path, autoCheck: update.autoCheck },
  }), [folder, template, output, periodMode, month, year, fromDate, toDate, force, ollama, update.path, update.autoCheck])

  const applyConfig = useCallback((config: NativeConfig) => {
    setFolderValue(config.paths?.reportFolder || '')
    setTemplateValue(config.paths?.template || '')
    setOutputValue(config.paths?.output || '')
    setPeriodModeValue(config.period?.mode || 'auto')
    setMonthValue(String(config.period?.month || ''))
    setYearValue(String(config.period?.year || ''))
    setFromDateValue(config.period?.from || '')
    setToDateValue(config.period?.to || '')
    setForceValue(Boolean(config.forceReprocess))
    setOllamaState({
      server: config.ollama?.host || '', port: String(config.ollama?.port || '11434'),
      model: config.ollama?.model || '', checked: config.ollama?.checked || 'unchecked',
      models: config.ollama?.models || [], message: config.ollama?.message || 'Chưa kiểm tra kết nối.',
    })
    setUpdateState((prev) => ({ ...prev, ...(config.update || {}) }))
    if (config.app) setAppInfo({ name: config.app.name, title: config.app.title, version: config.app.version, build: config.app.build })
  }, [])

  const acceptDashboard = useCallback((raw: DashboardDTO) => {
    if (!alive.current) return
    const data = normalizeDashboard(raw)
    setConnected(true)
    setAppInfo({ name: data.app.name, title: data.app.title, version: data.app.version, build: data.app.build })
    setReports(data.reports)
    setScanned(Boolean(data.scan.scanned))
    setScanMessage(data.scan.message || '')
    setRun(data.job)
    if (!ollamaDirty.current) {
      setOllamaState({
        server: data.ollama.host || '', port: String(data.ollama.port || '11434'), model: data.ollama.model || '',
        checked: data.ollama.checked, models: data.ollama.models || [], message: data.ollama.message || '',
      })
    } else {
      setOllamaState((current) => ({ ...current, checked: data.ollama.checked,
        models: data.ollama.models || current.models, message: data.ollama.message || current.message }))
    }
    setAiStatusText(data.ollama.aiStatus || '')
    setDiscovering(Boolean(data.ollama.discovering))
    setDiscoveryMessage(data.ollama.discoveryMessage || '')
    setDiscoveryResults(data.ollama.discoveryResults || [])
    setServerApplyRunning(Boolean(data.ollama.serverApplyRunning))
    setServerApplyMessage(data.ollama.serverApplyMessage || '')
    if (!updateDirty.current) setUpdateState(data.update)
    else setUpdateState((current) => ({ ...data.update, path: current.path, autoCheck: current.autoCheck }))
    setDiagRows(data.diagnostics.rows || [])
    setDiagRunning(Boolean(data.diagnostics.running))
    const freshLogs = (data.logs || []).filter((entry) => (entry.id || 0) > clearLogThroughRef.current).slice(-500)
    setLog(freshLogs)
    setErrorMessage((current) => current.includes('Python bridge is not connected') ? '' : current)
    setErrorCode((current) => current === 'BRIDGE_UNAVAILABLE' ? '' : current)
    if (data.reports.length === 0) {
      setActiveIdValue('')
      setSelectedIds([])
    } else {
      setActiveIdValue((current) => data.reports.some((r) => r.id === current) ? current : data.reports[0].id)
    }
    if (data.job.error) { setErrorCode('WORKER_FAILED'); setErrorMessage(data.job.error) }
  }, [])

  const refreshDashboard = useCallback(async () => {
    if (dashboardInFlight.current) return
    dashboardInFlight.current = true
    try {
      const raw = await callBridge<DashboardDTO>('get_dashboard_state')
      acceptDashboard(raw)
    } catch (error) {
      if (alive.current) {
        setConnected(false)
        recordError(error)
      }
    } finally {
      dashboardInFlight.current = false
    }
  }, [acceptDashboard, recordError])

  const refreshLearning = useCallback(async () => {
    if (learningInFlight.current || !hasNativeBridge()) return
    learningInFlight.current = true
    lastLearningRefresh.current = Date.now()
    try {
      const raw = await callBridge<LearningState & { training?: boolean }>('get_learning_state')
      if (!alive.current) return
      setLearningState({ ...EMPTY_LEARNING, ...raw })
      setTraining(Boolean(raw.training))
    } catch (error) {
      if (alive.current) recordError(error)
    } finally {
      learningInFlight.current = false
    }
  }, [recordError])

  useEffect(() => {
    alive.current = true
    let cancelled = false
    const initialize = async () => {
      try {
        const [version, config] = await Promise.all([
          callBridge<typeof appInfo>('get_app_version'),
          callBridge<NativeConfig>('get_current_config'),
        ])
        if (cancelled) return
        setConnected(true)
        setAppInfo({ name: version.name, title: version.title, version: version.version, build: version.build })
        applyConfig(config)
        await refreshDashboard()
      } catch (error) {
        if (cancelled) return
        setConnected(false)
        recordError(error)
      }
    }
    void initialize()
    const timer = window.setInterval(() => {
      void refreshDashboard()
      if (tabRef.current === 'learning' && Date.now() - lastLearningRefresh.current > 1400) void refreshLearning()
    }, 750)
    const onDeferredClose = (event: Event) => {
      const detail = (event as CustomEvent<string>).detail
      if (detail) { setErrorCode('BUSY'); setErrorMessage(detail) }
    }
    window.addEventListener('tnp-close-requested', onDeferredClose)
    return () => {
      cancelled = true
      alive.current = false
      window.clearInterval(timer)
      window.removeEventListener('tnp-close-requested', onDeferredClose)
    }
  }, [applyConfig, recordError, refreshDashboard, refreshLearning])

  const setTab = useCallback((value: TabKey) => {
    tabRef.current = value
    setTabValue(value)
    if (value === 'learning') void refreshLearning()
  }, [refreshLearning])

  const setActiveId = useCallback((id: string) => setActiveIdValue(id), [])
  const activeReport = reports.find((report) => report.id === activeId) ?? null

  const fetchReportDetails = useCallback(async (report: Report) => {
    try {
      const details = await callBridge<{
        causes: string[]; countermeasures: string[]; afterImages: Report['results']['afterImages'];
        qpn?: Report['results']['qpn']; slides: number; excelRow?: number; defectText?: string;
        temporaryRemoved?: boolean; warning?: string; occurrenceDate?: string; vendor?: string;
      }>('get_report_details', report.id)
      if (!alive.current) return
      setReports((current) => current.map((item) => item.id !== report.id ? item : ({
        ...item,
        slides: details.slides || item.slides,
        excelRow: details.excelRow ?? item.excelRow,
        occurrenceDate: details.occurrenceDate || item.occurrenceDate,
        vendor: details.vendor || item.vendor,
        warning: details.warning || item.warning,
        results: { ...item.results, qpn: details.qpn, causes: details.causes || [],
          countermeasures: details.countermeasures || [], afterImages: details.afterImages || [],
          temporaryRemoved: Boolean(details.temporaryRemoved) },
      })))
      setReveal((current) => ({ ...current, [report.id]: REVEAL_DONE }))
    } catch (error) {
      if (alive.current) recordError(error)
    }
  }, [recordError])

  useEffect(() => {
    if (!activeReport || !connected || ['waiting', 'new_row', 'processing'].includes(activeReport.status)) return
    const signature = `${activeReport.status}:${activeReport.shortResult}:${activeReport.excelRow || ''}`
    if (detailStatus.current[activeReport.id] === signature) return
    detailStatus.current[activeReport.id] = signature
    void fetchReportDetails(activeReport)
  }, [activeReport, connected, fetchReportDetails])

  const stats = useMemo(() => {
    const counts = { completed: 0, needs_review: 0, error: 0, skipped: 0, total: reports.length }
    for (const report of reports) {
      const bucket = BUCKET_OF[report.status]
      if (bucket !== 'other') counts[bucket] += 1
    }
    return counts
  }, [reports])

  const running = run.status === 'processing' || run.status === 'stopping'
  const locked = running || scanBusy || training || diagRunning || serverApplyRunning
  const queue = useMemo(() => {
    if (running) return run.queue.slice(run.doneCount)
    if (run.status === 'done' && !run.stopped) return []
    return reports.filter((report) => report.status === 'waiting' || report.status === 'new_row').map((report) => report.id)
  }, [reports, run, running])
  const filteredReports = useMemo(() => reports.filter((report) => {
    const haystack = `${report.managementNumber} ${report.fileName}`.toLowerCase()
    if (search.trim() && !haystack.includes(search.trim().toLowerCase())) return false
    switch (filter) {
      case 'File cần xử lý': return ['waiting', 'new_row', 'processing'].includes(report.status)
      case 'File đã xử lý': return report.status === 'completed'
      case 'File cần kiểm tra': return report.status === 'needs_review'
      case 'File lỗi': return report.status === 'error'
      case 'File bị bỏ qua': return ['skipped', 'outside_period', 'source_duplicate', 'fast_skip'].includes(report.status)
      case 'File đã loại thủ công': return report.status === 'excluded'
      default: return true
    }
  }), [reports, search, filter])

  const runOperation = useCallback(async (operation: () => Promise<unknown>) => {
    clearError()
    try {
      await operation()
      await refreshDashboard()
    } catch (error) {
      recordError(error)
    }
  }, [clearError, recordError, refreshDashboard])

  const dirtySource = useCallback(() => { setScanDirty(true); setScanned(false) }, [])
  const setFolder = useCallback((value: string) => {
    setFolderValue(value)
    setInputFileToken('')
    setInputFileName('')
    dirtySource()
  }, [dirtySource])
  const setTemplate = useCallback((value: string) => { setTemplateValue(value); dirtySource() }, [dirtySource])
  const setOutput = useCallback((value: string) => { setOutputValue(value); dirtySource() }, [dirtySource])
  const setPeriodMode = useCallback((value: PeriodMode) => { setPeriodModeValue(value); dirtySource() }, [dirtySource])
  const setMonth = useCallback((value: string) => { setMonthValue(value); dirtySource() }, [dirtySource])
  const setYear = useCallback((value: string) => { setYearValue(value); dirtySource() }, [dirtySource])
  const setFromDate = useCallback((value: string) => { setFromDateValue(value); dirtySource() }, [dirtySource])
  const setToDate = useCallback((value: string) => { setToDateValue(value); dirtySource() }, [dirtySource])
  const setForce = useCallback((value: boolean) => { setForceValue(value); dirtySource() }, [dirtySource])
  const setOllama = useCallback((value: OllamaConnection) => {
    ollamaDirty.current = true
    setOllamaState(value)
  }, [])
  const setUpdate = useCallback((value: UpdateState) => {
    updateDirty.current = true
    setUpdateState(value)
  }, [])

  const scan = useCallback(async () => {
    if (running) return
    setScanBusy(true)
    clearError()
    setScanMessage('Đang quét danh sách bằng bộ quét Python…')
    try {
      const raw = await callBridge<DashboardDTO>('scan_reports', configDto(), inputFileToken)
      detailStatus.current = {}
      ollamaDirty.current = false
      updateDirty.current = false
      setSelectedIds([])
      setReveal({})
      acceptDashboard(raw)
      setScanDirty(false)
    } catch (error) {
      recordError(error)
    } finally {
      setScanBusy(false)
    }
  }, [acceptDashboard, clearError, configDto, inputFileToken, recordError, running])

  const chooseReportFolder = useCallback(async () => {
    clearError()
    try {
      const result = await callBridge<{ cancelled: boolean; path?: string }>('choose_report_folder')
      if (!result.cancelled && result.path) setFolder(result.path)
    } catch (error) { recordError(error) }
  }, [clearError, recordError, setFolder])

  const choosePptxFile = useCallback(async () => {
    clearError()
    try {
      const result = await callBridge<{ cancelled: boolean; token?: string; name?: string; folder?: string }>('choose_pptx_file')
      if (!result.cancelled && result.token) {
        setInputFileToken(result.token)
        setInputFileName(result.name || '')
        if (result.folder) setFolderValue(result.folder)
        dirtySource()
      }
    } catch (error) { recordError(error) }
  }, [clearError, dirtySource, recordError])

  const chooseTemplateFile = useCallback(async () => {
    clearError()
    try {
      const result = await callBridge<{ cancelled: boolean; path?: string }>('choose_template_file')
      if (!result.cancelled && result.path) setTemplate(result.path)
    } catch (error) { recordError(error) }
  }, [clearError, recordError, setTemplate])

  const chooseOutputFile = useCallback(async () => {
    clearError()
    try {
      const result = await callBridge<{ cancelled: boolean; path?: string }>('choose_output_file')
      if (!result.cancelled && result.path) setOutput(result.path)
    } catch (error) { recordError(error) }
  }, [clearError, recordError, setOutput])

  const chooseUpdateFolder = useCallback(async () => {
    clearError()
    try {
      const result = await callBridge<{ cancelled: boolean; path?: string }>('choose_update_folder')
      if (!result.cancelled && result.path) {
        updateDirty.current = true
        setUpdateState((prev) => ({ ...prev, path: result.path || '' }))
      }
    } catch (error) { recordError(error) }
  }, [clearError, recordError])

  const changeSelected = useCallback(async (restore: boolean) => {
    const ids = selectedIds.length ? selectedIds : activeId ? [activeId] : []
    if (!ids.length) return
    await runOperation(() => callBridge(restore ? 'restore_reports' : 'exclude_reports', ids))
    setSelectedIds([])
  }, [activeId, runOperation, selectedIds])
  const excludeSelected = useCallback(() => changeSelected(false), [changeSelected])
  const restoreSelected = useCallback(() => changeSelected(true), [changeSelected])
  const toggleSelected = useCallback((id: string) => {
    setSelectedIds((current) => current.includes(id) ? current.filter((entry) => entry !== id) : [...current, id])
  }, [])
  const toggleAllSelected = useCallback((ids: string[]) => {
    setSelectedIds((current) => ids.every((id) => current.includes(id)) ? current.filter((id) => !ids.includes(id)) : ids)
  }, [])

  const startProcessing = useCallback(async () => {
    if (running || scanDirty || !scanned) {
      if (scanDirty || !scanned) { setErrorCode('SCAN_REQUIRED'); setErrorMessage('Nguồn hoặc tùy chọn đã thay đổi. Quét lại trước khi xử lý.') }
      return
    }
    await runOperation(async () => {
      await callBridge('start_processing', configDto(), true)
      ollamaDirty.current = false
      updateDirty.current = false
    })
  }, [configDto, runOperation, running, scanDirty, scanned])

  const stopAfterCurrent = useCallback(() => runOperation(() => callBridge('stop_after_current')), [runOperation])
  const resetRun = clearError

  const saveConfiguration = useCallback(() => runOperation(async () => {
    await callBridge('save_configuration', configDto())
    ollamaDirty.current = false
    updateDirty.current = false
  }), [configDto, runOperation])
  const checkConnection = useCallback(() => runOperation(async () => {
    await callBridge('check_ollama_connection', configDto())
    ollamaDirty.current = false
    updateDirty.current = false
  }), [configDto, runOperation])
  const refreshModels = useCallback(() => runOperation(async () => {
    await callBridge('refresh_ollama_models', configDto())
    ollamaDirty.current = false
    updateDirty.current = false
  }), [configDto, runOperation])
  const saveOllamaConfig = saveConfiguration
  const startDiscovery = useCallback(() => runOperation(() => callBridge('start_ollama_discovery')), [runOperation])
  const stopDiscovery = useCallback(() => runOperation(() => callBridge('stop_ollama_discovery')), [runOperation])
  const useServer = useCallback((host: string, port: number) => runOperation(async () => {
    await callBridge('use_discovered_server', host, port)
    ollamaDirty.current = false
    setAppliedServer(`${host}:${port}`)
  }), [runOperation])

  const checkUpdate = useCallback((_startup = false) => runOperation(async () => {
    await callBridge('check_update', configDto())
    updateDirty.current = false
    ollamaDirty.current = false
  }), [configDto, runOperation])
  const installUpdate = useCallback(() => runOperation(async () => {
    const result = await callBridge<{ status: string; message: string }>('install_update')
    setUpdateState((current) => ({ ...current, status: 'error', message: result.message }))
  }), [runOperation])

  const runDiagnostics = useCallback(() => runOperation(() => callBridge('run_diagnostics')), [runOperation])
  const clearLogView = useCallback(() => {
    const through = log.reduce((latest, entry) => Math.max(latest, entry.id || 0), clearLogThroughRef.current)
    clearLogThroughRef.current = through
    setLog([])
  }, [log])
  const openOutputFile = useCallback(() => runOperation(() => callBridge('open_output_file')), [runOperation])
  const openOutputFolder = useCallback(() => runOperation(() => callBridge('open_output_folder')), [runOperation])
  const openLogFile = useCallback(() => runOperation(() => callBridge('open_log_file')), [runOperation])
  const openLogFolder = useCallback(() => runOperation(() => callBridge('open_log_folder')), [runOperation])
  const openLearningFolder = useCallback(() => runOperation(() => callBridge('open_learning_folder')), [runOperation])

  const setActiveManualFields = useCallback(async (id: string, fields: { vendor?: string; occurrence_date?: string }) => {
    clearError()
    try {
      const result = await callBridge<{ status: string; message: string; state: DashboardDTO }>('save_manual_fields', id, fields)
      if (result.state) acceptDashboard(result.state)
      setErrorCode(result.status === 'locked' ? 'EXCEL_LOCKED' : result.status === 'error' ? 'MANUAL_FIELDS_FAILED' : '')
      setErrorMessage(result.message)
      detailStatus.current[id] = ''
    } catch (error) { recordError(error) }
  }, [acceptDashboard, clearError, recordError])
  const retryManualFields = useCallback(async (id: string) => {
    clearError()
    try {
      const result = await callBridge<{ status: string; message: string; state: DashboardDTO }>('retry_manual_fields', id)
      if (result.state) acceptDashboard(result.state)
      setErrorCode(result.status === 'locked' ? 'EXCEL_LOCKED' : result.status === 'error' ? 'MANUAL_FIELDS_FAILED' : '')
      setErrorMessage(result.message)
      detailStatus.current[id] = ''
    } catch (error) { recordError(error) }
  }, [acceptDashboard, clearError, recordError])

  const labelImage = useCallback(async (id: string, label: string) => {
    const candidate = learningState.images.find((item) => item.id === id)
    try {
      const next = await callBridge<LearningState>('set_learning_label', 'image', id, label, candidate?.note || '')
      setLearningState(next)
    } catch (error) { recordError(error) }
  }, [learningState.images, recordError])
  const labelContent = useCallback(async (id: string, label: string) => {
    const candidate = learningState.contents.find((item) => item.id === id)
    try {
      const next = await callBridge<LearningState>('set_learning_label', 'content', id, label, candidate?.note || '')
      setLearningState(next)
    } catch (error) { recordError(error) }
  }, [learningState.contents, recordError])
  const setNote = useCallback((kind: 'image' | 'content', id: string, note: string) => {
    setLearningState((current) => kind === 'image'
      ? { ...current, images: current.images.map((item) => item.id === id ? { ...item, note } : item) }
      : { ...current, contents: current.contents.map((item) => item.id === id ? { ...item, note } : item) })
  }, [])
  const saveLearningLabels = useCallback(async (kind: 'image' | 'content') => {
    const notes = kind === 'image'
      ? Object.fromEntries(learningState.images.map((candidate) => [candidate.id, candidate.note]))
      : Object.fromEntries(learningState.contents.map((candidate) => [candidate.id, candidate.note]))
    try {
      const result = await callBridge<{ message: string; state: LearningState }>('save_learning_labels', kind, notes)
      setLearningState({ ...EMPTY_LEARNING, ...result.state })
      setErrorCode(result.state.excelLastResult?.kind === 'locked' ? 'EXCEL_LOCKED' : '')
      setErrorMessage(result.message)
      setTraining(false)
    } catch (error) { recordError(error) }
  }, [learningState, recordError])
  const saveImageLabels = useCallback(() => saveLearningLabels('image'), [saveLearningLabels])
  const saveContentLabels = useCallback(() => saveLearningLabels('content'), [saveLearningLabels])
  const saveLearningCandidate = useCallback(async (kind: 'image' | 'content', id: string, label: string, note: string) => {
    try {
      await callBridge<LearningState>('set_learning_label', kind, id, label, note)
      const result = await callBridge<{ message: string; state: LearningState }>('save_learning_labels', kind, { [id]: note })
      setLearningState({ ...EMPTY_LEARNING, ...result.state })
      setErrorCode(result.state.excelLastResult?.kind === 'locked' ? 'EXCEL_LOCKED' : '')
      setErrorMessage(result.message)
    } catch (error) { recordError(error) }
  }, [recordError])
  const saveImageCandidate = useCallback((id: string, label: string, note: string) =>
    saveLearningCandidate('image', id, label, note), [saveLearningCandidate])
  const saveContentCandidate = useCallback((id: string, label: string, note: string) =>
    saveLearningCandidate('content', id, label, note), [saveLearningCandidate])
  const trainModels = useCallback(async () => {
    try {
      const result = await callBridge<LearningState & { training?: boolean }>('train_models')
      setLearningState({ ...EMPTY_LEARNING, ...result })
      setTraining(Boolean(result.training))
    } catch (error) { recordError(error) }
  }, [recordError])
  const exportLearningData = useCallback(async () => {
    clearError()
    try {
      const result = await callBridge<{ cancelled?: boolean; message?: string }>('export_learning_data')
      if (result.message) setErrorMessage(result.message)
    } catch (error) { recordError(error) }
  }, [clearError, recordError])
  const applyLabelsToExcel = useCallback(async (_retry = false) => {
    try {
      const result = await callBridge<{ message: string; state: LearningState }>('retry_learning_excel')
      setLearningState({ ...EMPTY_LEARNING, ...result.state })
      setErrorCode(result.state.excelLastResult?.kind === 'locked' ? 'EXCEL_LOCKED' : '')
      setErrorMessage(result.message)
    } catch (error) { recordError(error) }
  }, [recordError])
  const dismissExcelResult = useCallback(async () => {
    try {
      const next = await callBridge<LearningState>('dismiss_learning_excel_notice')
      setLearningState({ ...EMPTY_LEARNING, ...next })
    } catch (error) { recordError(error) }
  }, [recordError])

  const title = appInfo.title || appInfo.name
  const aiMode: 'model' | 'heuristic' = ollama.checked === 'ok' ? 'model' : 'heuristic'
  const ollamaIndicator = useMemo(() => {
    if (ollama.checked === 'ok') return { text: `Ollama: ${ollama.model} — Sẵn sàng`, tone: 'ok' as const }
    if (ollama.checked === 'fail') return { text: `Ollama: ${ollama.model} — không kết nối (dự phòng theo quy tắc)`, tone: 'warn' as const }
    return { text: `Ollama: ${ollama.model || 'chưa chọn'} — chưa kiểm tra`, tone: 'muted' as const }
  }, [ollama])

  const store = {
    tab: tabValue, setTab,
    connected, errorMessage, errorCode, clearError, appInfo, title,
    reports, filteredReports, stats, queue, scanned, scanDirty, scanBusy, scanMessage,
    activeId, setActiveId, activeReport, selectedIds, toggleSelected, toggleAllSelected,
    search, setSearch, filter, setFilter, scan, excludeSelected, restoreSelected,
    folder, setFolder, template, setTemplate, output, setOutput,
    periodMode, setPeriodMode, month, setMonth, year, setYear, fromDate, setFromDate, toDate, setToDate,
    force, setForce, run, reveal, running, locked, startProcessing, stopAfterCurrent, resetRun,
    log, logCollapsed, setLogCollapsed, clearLogView, openLogFile, openLogFolder, openOutputFile, openOutputFolder,
    ollama, setOllama, ollamaIndicator, aiStatusText, aiMode, discovering, discoveryMessage, discoveryResults,
    startDiscovery, stopDiscovery, useServer, appliedServer, serverApplyRunning, serverApplyMessage,
    checkConnection, refreshModels, saveOllamaConfig,
    update, setUpdate, checkUpdate, installUpdate, chooseUpdateFolder,
    diagRows, diagRunning, runDiagnostics,
    learning: learningState, refreshLearning, imageReviewIndex, setImageReviewIndex, contentReviewIndex, setContentReviewIndex,
    labelImage, labelContent, setNote, saveImageLabels, saveContentLabels, saveImageCandidate, saveContentCandidate,
    trainModels, exportLearningData,
    openLearningFolder, applyLabelsToExcel, dismissExcelResult, training,
    inputFileName, chooseReportFolder, choosePptxFile, chooseTemplateFile, chooseOutputFile,
    saveManualFields: setActiveManualFields, retryManualFields,
  }
  return store
}

export function StoreProvider({ children }: { children: ReactNode }) {
  const value = useStoreValue()
  return <StoreContext.Provider value={value}>{children}</StoreContext.Provider>
}

export function useStore(): StoreValue {
  const ctx = useContext(StoreContext)
  if (!ctx) throw new Error('useStore must be used inside StoreProvider')
  return ctx
}
