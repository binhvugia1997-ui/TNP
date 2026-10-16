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
import {
  BridgeCallError,
  BridgeMethodMissingError,
  BridgeProtocolError,
  BridgeUnavailableError,
  callBridgeWithOptions,
  isAbortError,
  isBridgeConnectionError,
  resetBridgeHandshake,
} from '../lib/bridge'
import { createBridgeSessionController, startAppBridge, type BridgeSessionController } from '../lib/startup'
import {
  BUCKET_OF,
  presentReports,
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
  'File đã loại thủ công',
] as const
export type FilterKey = (typeof FILTERS)[number]
export type PeriodMode = 'auto' | 'all' | 'month' | 'range'

export type UpdateStatus = 'idle' | 'checking' | 'available' | 'latest' | 'older' | 'error'
  | 'development_disabled' | 'staging' | 'verifying' | 'installing' | 'ready_to_restart'
  | 'failed' | 'installed'

export type UpdateState = {
  path: string
  autoCheck: boolean
  status: UpdateStatus
  message: string
  available: { version: string; build: string; package: string; size: number | null } | null
  progress: number
  progressStage: string
  installEnabled: boolean
  canInstall: boolean
  installDisabledReason: string
  error: string
}

export type LearningCandidateAvailability = {
  state: 'ready' | 'processing' | 'no_run' | 'unavailable'
  reason: 'available' | 'empty' | 'processing' | 'no_run' | 'unavailable'
  count: number
  message: string
}

export type LearningState = {
  images: ImageCandidate[]
  contents: ContentCandidate[]
  candidateStatus: { image: LearningCandidateAvailability; content: LearningCandidateAvailability }
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
  app?: { name: string; title: string; version: string; build: string; buildNumber: number;
    runtime?: string; updateInstallCapable?: boolean }
}

type AppInfo = { name: string; title: string; version: string; build: string; buildNumber: number;
  runtime?: string; updateInstallCapable?: boolean }

type DashboardDTO = {
  app: AppInfo
  config: NativeConfig
  scan: {
    scanned: boolean
    message: string
    generation?: number
    revision?: number
    inProgress?: boolean
    incremental?: boolean
    completedLeaves?: number
    totalLeaves?: number
    firstLeafMs?: number | null
  }
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
  update: unknown
  diagnostics: { running: boolean; rows: DiagnosticRow[]; error: string }
}

type NormalizedDashboardDTO = Omit<DashboardDTO, 'update'> & { update: UpdateState }

export type ScanSnapshotVersion = { generation: number; revision: number }

function scanSnapshotVersion(scan: DashboardDTO['scan']): ScanSnapshotVersion {
  return {
    generation: Number.isFinite(scan?.generation) ? Number(scan.generation) : 0,
    revision: Number.isFinite(scan?.revision) ? Number(scan.revision) : 0,
  }
}

/** Reject delayed poll responses and the previous completed generation while a new Quét is pending. */
export function shouldAcceptScanSnapshot(
  current: ScanSnapshotVersion,
  incoming: DashboardDTO['scan'],
  pendingBaselineGeneration: number | null = null,
): boolean {
  const next = scanSnapshotVersion(incoming)
  if (next.generation < current.generation) return false
  if (next.generation === current.generation && next.revision < current.revision) return false
  if (pendingBaselineGeneration !== null && next.generation <= pendingBaselineGeneration && !incoming.inProgress) {
    return false
  }
  return true
}

function mergeIncrementalReports(current: Report[], incoming: Report[]): Report[] {
  const byId = new Map(current.map((item) => [item.id, item]))
  // PROMPT-037: the delta copy always overwrites the previous one, even when it is now hidden
  // (skipped). A visible → skipped transition therefore drops the stale visible row below.
  for (const item of incoming) byId.set(item.id, item)
  const merged = [...byId.values()].sort((left, right) => {
    const a = Number(left.id)
    const b = Number(right.id)
    return Number.isFinite(a) && Number.isFinite(b) ? a - b : left.id.localeCompare(right.id)
  })
  return presentReports(merged)
}

/**
 * PROMPT-037: the active visible row after a reconciliation. Keeps it while visible; when it
 * became hidden, advances to the next visible row after its previous position, else the previous
 * visible row, else the first visible row. Empty list → ''. O(N).
 */
export function resolveActiveReportId(
  previous: readonly Report[],
  next: readonly Report[],
  activeId: string,
  nextIds: ReadonlySet<string> = new Set(next.map((report) => report.id)),
): string {
  if (next.length === 0) return ''
  if (nextIds.has(activeId)) return activeId
  const at = previous.findIndex((report) => report.id === activeId)
  if (at >= 0) {
    for (let i = at + 1; i < previous.length; i += 1) if (nextIds.has(previous[i].id)) return previous[i].id
    for (let i = at - 1; i >= 0; i -= 1) if (nextIds.has(previous[i].id)) return previous[i].id
  }
  return next[0].id
}

const EMPTY_RUN: ProcessingRun = {
  status: 'idle', queue: [], index: 0, doneCount: 0, stage: 'waiting', percent: 0, currentFile: '',
  elapsedSec: 0, remainSec: null, startedAt: null, finishedAt: null, hasSamples: false,
}
const EMPTY_UPDATE: UpdateState = {
  path: '', autoCheck: true, status: 'idle', message: 'Chưa kiểm tra cập nhật.', available: null,
  progress: 0, progressStage: '', installEnabled: false, canInstall: false,
  installDisabledReason: '', error: '',
}
const UPDATE_STATUSES = new Set<UpdateStatus>([
  'idle', 'checking', 'available', 'latest', 'older', 'error', 'development_disabled',
  'staging', 'verifying', 'installing', 'ready_to_restart', 'failed', 'installed',
])

const EMPTY_LEARNING: LearningState = {
  images: [], contents: [],
  candidateStatus: {
    image: { state: 'no_run', reason: 'no_run', count: 0, message: 'Chưa tải trạng thái ảnh ứng viên.' },
    content: { state: 'no_run', reason: 'no_run', count: 0, message: 'Chưa tải trạng thái nội dung.' },
  },
  counts: { image: { total: 0, labeled: 0 }, content: { total: 0, labeled: 0 } },
  modelStatus: { image: 'Chưa tải trạng thái mô hình.', content: 'Chưa tải trạng thái mô hình.' },
  excelPending: 0,
  excelLastResult: { kind: 'none', message: '' },
}

export type LearningLoadState = { status: 'idle' | 'loading' | 'ready' | 'error'; message: string }

/** PROMPT-036S: minimal bridge acks for the label fast paths (never a full LearningState). */
export type LearningLabelAck = {
  kind: 'image' | 'content'
  candidateId: string
  label: string
  labelPending: boolean
}
export type LearningSaveAck = LearningLabelAck & {
  message: string
  excelPending: number
  excelLastResult: { kind: 'ok' | 'locked' | 'none'; message: string }
}

type LocalLabelEntry = {
  kind: 'image' | 'content'
  id: string
  label: string
  note: string
  phase: 'pending' | 'acked' | 'saving' | 'error'
}

function normalizeLearning(raw: unknown): LearningState {
  if (!raw || typeof raw !== 'object') throw new BridgeProtocolError()
  const value = raw as Partial<LearningState>
  if (!Array.isArray(value.images) || !Array.isArray(value.contents)) throw new BridgeProtocolError()
  const states = new Set(['ready', 'processing', 'no_run', 'unavailable'])
  const reasons = new Set(['available', 'empty', 'processing', 'no_run', 'unavailable'])
  const image = value.candidateStatus?.image
  const content = value.candidateStatus?.content
  for (const [status, candidates] of [[image, value.images], [content, value.contents]] as const) {
    if (!status || !states.has(status.state) || !reasons.has(status.reason)
      || !Number.isInteger(status.count) || status.count < 0 || typeof status.message !== 'string'
      || status.count !== candidates.length) {
      throw new BridgeProtocolError()
    }
  }
  if (!value.counts?.image || !value.counts?.content || !value.modelStatus
    || typeof value.modelStatus.image !== 'string' || typeof value.modelStatus.content !== 'string') {
    throw new BridgeProtocolError()
  }
  return { ...EMPTY_LEARNING, ...value,
    images: value.images.map((candidate) => ({ ...candidate, id: candidate.reviewId || candidate.id })),
    contents: value.contents,
    candidateStatus: {
      image: image as LearningCandidateAvailability,
      content: content as LearningCandidateAvailability,
    } }
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

const INVALID_UPDATE_MESSAGE = 'Phản hồi trạng thái cập nhật từ Python bị thiếu hoặc không hợp lệ.'

function invalidUpdateState(message = INVALID_UPDATE_MESSAGE): UpdateState {
  return { ...EMPTY_UPDATE, status: 'failed', message, error: message,
    installDisabledReason: 'Không thể cài đặt khi trạng thái cập nhật không hợp lệ.' }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value)
}

/** Keep an incomplete or malformed native update DTO visible as a failure, never as install authorization. */
export function normalizeUpdate(raw: unknown): UpdateState {
  if (!isRecord(raw) || typeof raw.path !== 'string' || typeof raw.autoCheck !== 'boolean'
    || typeof raw.status !== 'string' || !UPDATE_STATUSES.has(raw.status as UpdateStatus)
    || typeof raw.message !== 'string' || typeof raw.progress !== 'number'
    || !Number.isFinite(raw.progress) || raw.progress < 0 || raw.progress > 100
    || typeof raw.progressStage !== 'string' || typeof raw.installEnabled !== 'boolean'
    || typeof raw.canInstall !== 'boolean' || typeof raw.installDisabledReason !== 'string'
    || typeof raw.error !== 'string') return invalidUpdateState()

  let available: UpdateState['available'] = null
  if (raw.available !== null) {
    if (!isRecord(raw.available) || typeof raw.available.version !== 'string'
      || typeof raw.available.build !== 'string' || !/^\d{1,12}$/.test(raw.available.build)
      || typeof raw.available.package !== 'string' || raw.available.package.length === 0
      || !(raw.available.size === null || (typeof raw.available.size === 'number'
        && Number.isSafeInteger(raw.available.size) && raw.available.size >= 0))) return invalidUpdateState()
    available = { version: raw.available.version, build: raw.available.build,
      package: raw.available.package, size: raw.available.size as number | null }
  }

  const normalized: UpdateState = {
    path: raw.path, autoCheck: raw.autoCheck, status: raw.status as UpdateStatus,
    message: raw.message, available, progress: raw.progress, progressStage: raw.progressStage,
    installEnabled: raw.installEnabled, canInstall: raw.canInstall,
    installDisabledReason: raw.installDisabledReason, error: raw.error,
  }
  if ((normalized.status === 'available' && !available)
    || (normalized.installEnabled && (normalized.status !== 'available' || !normalized.canInstall))) {
    return invalidUpdateState()
  }
  return normalized
}

function normalizeDashboard(raw: DashboardDTO): NormalizedDashboardDTO {
  const reports = (raw.reports || []).map(reportFromDto)
  return {
    ...raw,
    update: normalizeUpdate(raw.update),
    scan: { ...raw.scan, generation: raw.scan?.generation ?? 0, revision: raw.scan?.revision ?? 0,
      inProgress: raw.scan?.inProgress ?? false, incremental: raw.scan?.incremental ?? false,
      completedLeaves: raw.scan?.completedLeaves ?? 0,
      totalLeaves: raw.scan?.totalLeaves ?? 0, firstLeafMs: raw.scan?.firstLeafMs ?? null },
    reports,
    job: { ...EMPTY_RUN, ...(raw.job || {}), stage: stageKey(raw.job?.stage) },
    ollama: { ...EMPTY_OLLAMA, ...(raw.ollama || {}) },
    diagnostics: raw.diagnostics || { running: false, rows: [], error: '' },
    logs: raw.logs || [],
  }
}

function safeErrorDetails(error: unknown): { code: string; message: string } {
  if (error instanceof BridgeCallError || error instanceof BridgeUnavailableError
    || error instanceof BridgeMethodMissingError || error instanceof BridgeProtocolError) {
    return { code: error.code, message: error.message }
  }
  // Unclassified JavaScript exceptions can contain stack traces, local paths, or implementation details.
  return { code: 'INTERNAL_ERROR', message: 'Thao tác không thành công. Hãy kiểm tra nhật ký ứng dụng.' }
}

function useStoreValue() {
  const [tabValue, setTabValue] = useState<TabKey>('reports')
  const [connected, setConnected] = useState(false)
  const [errorMessage, setErrorMessage] = useState('')
  const [errorCode, setErrorCode] = useState('')
  const [appInfo, setAppInfo] = useState<AppInfo>({ name: 'Report Extractor', title: 'Report Extractor',
    version: '', build: '', buildNumber: 0, runtime: 'pywebview-development', updateInstallCapable: false })
  const [reports, setReports] = useState<Report[]>([])
  const [scanned, setScanned] = useState(false)
  const [scanDirty, setScanDirty] = useState(false)
  const [scanBusy, setScanBusy] = useState(false)
  const [scanInProgress, setScanInProgress] = useState(false)
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
  const [update, setUpdateState] = useState<UpdateState>(EMPTY_UPDATE)
  const [diagRows, setDiagRows] = useState<DiagnosticRow[]>([])
  const [diagRunning, setDiagRunning] = useState(false)
  const [learningState, setLearningState] = useState<LearningState>(EMPTY_LEARNING)
  const [learningLoad, setLearningLoad] = useState<LearningLoadState>({ status: 'idle', message: '' })
  const [imageReviewIndex, setImageReviewIndex] = useState(0)
  const [contentReviewIndex, setContentReviewIndex] = useState(0)
  const [training, setTraining] = useState(false)
  const [inputFileToken, setInputFileToken] = useState('')
  const [inputFileName, setInputFileName] = useState('')
  const dashboardInFlight = useRef(false)
  const reportsRef = useRef<Report[]>([])
  const scanBusyRef = useRef(false)
  const scanVersionRef = useRef<ScanSnapshotVersion>({ generation: 0, revision: 0 })
  const scanBaselineGenerationRef = useRef<number | null>(null)
  const learningInFlight = useRef(false)
  const lastLearningRefresh = useRef(0)
  // PROMPT-036S: local (optimistic) learning labels. The map is the source of truth for a
  // candidate whose newest user choice has not been confirmed saved by the backend; it is
  // merged over every incoming full LearningState so a stale poll can never revert a click.
  // Phases: 'pending' = fast-set not yet acked, 'acked' = backend holds the pending label,
  // 'saving' = atomic save in flight, 'error' = last persist failed (label stays visible,
  // marked unsaved via labelPending, and can be retried). Entries drop out when a full
  // state confirms the same label+note, or when a save ack lands.
  const localLabelsRef = useRef(new Map<string, LocalLabelEntry>())
  const labelWorkersRef = useRef(new Set<string>())
  const labelWritesInFlightRef = useRef(0)
  const labelWriteEpochRef = useRef(0)
  const learningRefreshQueuedRef = useRef(false)
  const learningStateRef = useRef<LearningState>(EMPTY_LEARNING)
  const bridgeReadyRef = useRef(false)
  const bridgeLifetimeRef = useRef<AbortController | null>(null)
  const bridgeSessionRef = useRef<BridgeSessionController | null>(null)
  const tabRef = useRef<TabKey>(tabValue)
  const detailStatus = useRef<Record<string, string>>({})
  const ollamaDirty = useRef(false)
  const updateDirty = useRef(false)
  const autoUpdateCloseRequested = useRef(false)
  const updateInstallInFlight = useRef(false)
  const updateCloseInFlight = useRef(false)
  const alive = useRef(true)
  useEffect(() => { reportsRef.current = reports }, [reports])
  useEffect(() => { learningStateRef.current = learningState }, [learningState])
  const publishError = useCallback((error: unknown) => {
    if (!alive.current || isAbortError(error)) return
    const details = safeErrorDetails(error)
    setErrorCode((current) => current === details.code ? current : details.code)
    setErrorMessage((current) => current === details.message ? current : details.message)
  }, [])
  const recordError = useCallback((error: unknown) => {
    if (isAbortError(error)) return
    if (isBridgeConnectionError(error) && bridgeSessionRef.current) {
      bridgeSessionRef.current.invalidate(error)
      return
    }
    publishError(error)
  }, [publishError])
  const clearError = useCallback(() => { setErrorCode(''); setErrorMessage('') }, [])

  const callNative = useCallback(<T,>(method: string, ...args: unknown[]) =>
    callBridgeWithOptions<T>(method, args, { signal: bridgeLifetimeRef.current?.signal }), [])

  // PROMPT-036S: merge unconfirmed local labels over an incoming full LearningState. A
  // candidate whose backend row already shows the same label+note is confirmed and its
  // overlay entry is dropped; any other overlay wins so stale responses cannot revert a
  // click. Labeled counts mirror the backend rule (saved ∪ pending): a locally labeled
  // candidate the backend still shows UNLABELED adds one.
  const mergeLocalLabels = useCallback((next: LearningState): LearningState => {
    if (localLabelsRef.current.size === 0) return next
    let imageDelta = 0
    let contentDelta = 0
    const apply = <T extends { id: string; userLabel: string; note: string }>(list: readonly T[], kind: 'image' | 'content'): T[] => list.map((candidate) => {
      const key = `${kind}:${candidate.id}`
      const entry = localLabelsRef.current.get(key)
      if (!entry) return candidate
      if (entry.phase !== 'error' && entry.phase !== 'saving'
        && candidate.userLabel === entry.label && candidate.note === entry.note) {
        localLabelsRef.current.delete(key)
        return candidate
      }
      if (candidate.userLabel === 'UNLABELED' && entry.label !== 'UNLABELED') {
        if (kind === 'image') imageDelta += 1
        else contentDelta += 1
      }
      return { ...candidate, userLabel: entry.label, note: entry.note, labelPending: true }
    })
    const merged: LearningState = {
      ...next,
      images: apply(next.images, 'image'),
      contents: apply(next.contents, 'content'),
    }
    if (imageDelta || contentDelta) {
      merged.counts = {
        image: { ...merged.counts.image, labeled: merged.counts.image.labeled + imageDelta },
        content: { ...merged.counts.content, labeled: merged.counts.content.labeled + contentDelta },
      }
    }
    return merged
  }, [])

  const acceptLearningState = useCallback((next: LearningState) => {
    setLearningState(mergeLocalLabels(next))
  }, [mergeLocalLabels])

  const labelWritesActive = useCallback(() => labelWritesInFlightRef.current > 0
    || [...localLabelsRef.current.values()].some((entry) => entry.phase === 'saving'), [])

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
    if (config.app) setAppInfo({ ...config.app,
      runtime: config.app.runtime || 'pywebview-development',
      updateInstallCapable: Boolean(config.app.updateInstallCapable) })
  }, [])

  const acceptDashboard = useCallback((raw: DashboardDTO) => {
    if (!alive.current) return
    const data = normalizeDashboard(raw)
    const pendingBaseline = scanBusyRef.current ? scanBaselineGenerationRef.current : null
    if (!shouldAcceptScanSnapshot(scanVersionRef.current, data.scan, pendingBaseline)) return
    scanVersionRef.current = scanSnapshotVersion(data.scan)
    setScanInProgress(Boolean(data.scan.inProgress))
    const previousReports = reportsRef.current
    // PROMPT-037: both paths yield visible rows only (single rule: presentReports / isVisibleReport).
    const nextReports = data.scan.incremental
      ? mergeIncrementalReports(previousReports, data.reports)
      : presentReports(data.reports)
    reportsRef.current = nextReports
    setAppInfo({ ...data.app, runtime: data.app.runtime || 'pywebview-development',
      updateInstallCapable: Boolean(data.app.updateInstallCapable) })
    setReports(nextReports)
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
    if (bridgeReadyRef.current) {
      setErrorMessage((current) => current.startsWith('Chưa kết nối ứng dụng Python.')
        || current.startsWith('Hợp đồng bridge Python') || current.startsWith('Python bridge trả về phản hồi') ? '' : current)
      setErrorCode((current) => ['BRIDGE_UNAVAILABLE', 'BRIDGE_METHOD_MISSING', 'BRIDGE_CONTRACT_ERROR'].includes(current) ? '' : current)
    }
    // PROMPT-037: selection and the active row only ever reference visible rows.
    const visibleIds = new Set(nextReports.map((report) => report.id))
    setSelectedIds((current) => {
      const kept = current.filter((id) => visibleIds.has(id))
      return kept.length === current.length ? current : kept
    })
    setActiveIdValue((current) => resolveActiveReportId(previousReports, nextReports, current, visibleIds))
    if (data.job.error) { setErrorCode('WORKER_FAILED'); setErrorMessage(data.job.error) }
  }, [])

  const loadDashboard = useCallback(async (signal?: AbortSignal) => {
    if (dashboardInFlight.current) return
    dashboardInFlight.current = true
    try {
      const raw = await callBridgeWithOptions<DashboardDTO>('get_dashboard_state', [], {
        signal: signal ?? bridgeLifetimeRef.current?.signal,
      })
      acceptDashboard(raw)
    } finally {
      dashboardInFlight.current = false
    }
  }, [acceptDashboard])
  const refreshDashboard = useCallback(() => loadDashboard(), [loadDashboard])

  const refreshLearning = useCallback(async (signal?: AbortSignal) => {
    if (learningInFlight.current || !bridgeReadyRef.current) return
    // PROMPT-036S: user label writes have priority over background refresh. Defer (and
    // re-queue) instead of racing a write; the write worker re-runs this when it drains.
    if (labelWritesActive()) {
      learningRefreshQueuedRef.current = true
      return
    }
    learningInFlight.current = true
    lastLearningRefresh.current = Date.now()
    const epochAtStart = labelWriteEpochRef.current
    if (alive.current) setLearningLoad({ status: 'loading', message: '' })
    try {
      const raw = await callBridgeWithOptions<unknown>('get_learning_state', [], {
        signal: signal ?? bridgeLifetimeRef.current?.signal,
      })
      if (!alive.current) return
      // A label changed while this response was in flight: it may predate the click, so
      // discard it and re-queue a fresh read instead of letting it revert the choice.
      if (epochAtStart !== labelWriteEpochRef.current) {
        learningRefreshQueuedRef.current = true
        return
      }
      const next = normalizeLearning(raw)
      acceptLearningState(next)
      setTraining(Boolean(next.training))
      setLearningLoad({ status: 'ready', message: '' })
    } catch (error) {
      if (alive.current && !isAbortError(error)) {
        recordError(error)
        setLearningLoad({ status: 'error', message: safeErrorDetails(error).message })
      }
    } finally {
      learningInFlight.current = false
      if (learningRefreshQueuedRef.current && alive.current && !labelWritesActive()) {
        learningRefreshQueuedRef.current = false
        void refreshLearning()
      }
    }
  }, [acceptLearningState, labelWritesActive, recordError])

  // Run a deferred Learning refresh once user label writes have drained.
  const drainLearningRefresh = useCallback(() => {
    if (!learningRefreshQueuedRef.current || !alive.current || labelWritesActive()) return
    learningRefreshQueuedRef.current = false
    void refreshLearning()
  }, [labelWritesActive, refreshLearning])

  // PROMPT-036S: per-candidate label persist worker. At most one in-flight
  // set_learning_label_fast per candidate; clicks that arrive during a flight only replace
  // the desired value, so rapid clicking collapses to at most one extra call and the last
  // choice always wins. The backlog is therefore bounded per candidate, never unbounded.
  const persistLabel = useCallback((kind: 'image' | 'content', id: string) => {
    const key = `${kind}:${id}`
    if (labelWorkersRef.current.has(key)) return
    labelWorkersRef.current.add(key)
    labelWritesInFlightRef.current += 1
    void (async () => {
      try {
        for (;;) {
          const entry = localLabelsRef.current.get(key)
          if (!entry || entry.phase === 'saving') return
          const sentLabel = entry.label
          const sentNote = entry.note
          entry.phase = 'pending'
          try {
            await callNative<LearningLabelAck>('set_learning_label_fast', kind, id, sentLabel, sentNote)
          } catch (error) {
            const current = localLabelsRef.current.get(key)
            if (current && current.phase !== 'saving'
              && (current.label !== sentLabel || current.note !== sentNote)) {
              continue // a newer click landed during the failed flight: persist that instead
            }
            if (current && current.phase !== 'saving') current.phase = 'error'
            if (alive.current) recordError(error)
            return
          }
          const current = localLabelsRef.current.get(key)
          if (!current || current.phase === 'saving') return
          current.phase = 'acked'
          if (current.label === sentLabel && current.note === sentNote) return
          // A newer click landed during the flight: loop and persist the latest value.
        }
      } finally {
        labelWorkersRef.current.delete(key)
        labelWritesInFlightRef.current -= 1
        drainLearningRefresh()
      }
    })()
  }, [callNative, drainLearningRefresh, recordError])

  const markBridgeDisconnected = useCallback((error: unknown) => {
    bridgeReadyRef.current = false
    resetBridgeHandshake()
    setConnected(false)
    publishError(error)
  }, [publishError])

  useEffect(() => {
    alive.current = true
    const lifetime = new AbortController()
    bridgeLifetimeRef.current = lifetime
    const onDeferredClose = (event: Event) => {
      const detail = (event as CustomEvent<string>).detail
      if (detail) { setErrorCode('BUSY'); setErrorMessage(detail) }
    }
    window.addEventListener('tnp-close-requested', onDeferredClose)

    const session = createBridgeSessionController({
      connect: (signal) => startAppBridge<typeof appInfo, NativeConfig, DashboardDTO>(signal),
      poll: async (signal) => {
        const tasks: Promise<void>[] = [loadDashboard(signal)]
        if (tabRef.current === 'learning' && Date.now() - lastLearningRefresh.current > 1400) {
          tasks.push(refreshLearning(signal))
        }
        await Promise.all(tasks)
      },
      onConnected: ({ version, config, dashboard }) => {
        bridgeReadyRef.current = true
        setAppInfo({ ...version, runtime: version.runtime || 'pywebview-development',
          updateInstallCapable: Boolean(version.updateInstallCapable) })
        applyConfig(config)
        acceptDashboard(dashboard)
        setConnected(true)
      },
      onDisconnected: markBridgeDisconnected,
      onPollError: recordError,
    })
    bridgeSessionRef.current = session
    session.start()

    return () => {
      alive.current = false
      bridgeReadyRef.current = false
      lifetime.abort()
      session.stop()
      resetBridgeHandshake()
      if (bridgeSessionRef.current === session) bridgeSessionRef.current = null
      if (bridgeLifetimeRef.current === lifetime) bridgeLifetimeRef.current = null
      window.removeEventListener('tnp-close-requested', onDeferredClose)
    }
  }, [acceptDashboard, applyConfig, loadDashboard, markBridgeDisconnected, recordError, refreshLearning])

  const setTab = useCallback((value: TabKey) => {
    tabRef.current = value
    setTabValue(value)
    if (value === 'learning') void refreshLearning()
  }, [refreshLearning])

  const setActiveId = useCallback((id: string) => setActiveIdValue(id), [])
  const activeReport = reports.find((report) => report.id === activeId) ?? null

  const fetchReportDetails = useCallback(async (report: Report) => {
    try {
      const details = await callNative<{
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
    if (!activeReport || !connected || scanBusy || scanInProgress
      || ['waiting', 'new_row', 'processing'].includes(activeReport.status)) return
    const signature = `${activeReport.status}:${activeReport.shortResult}:${activeReport.excelRow || ''}`
    if (detailStatus.current[activeReport.id] === signature) return
    detailStatus.current[activeReport.id] = signature
    void fetchReportDetails(activeReport)
  }, [activeReport, connected, fetchReportDetails, scanBusy, scanInProgress])

  const stats = useMemo(() => {
    const counts = { completed: 0, needs_review: 0, error: 0, skipped: 0, total: reports.length }
    for (const report of reports) {
      const bucket = BUCKET_OF[report.status]
      if (bucket !== 'other') counts[bucket] += 1
    }
    return counts
  }, [reports])

  // PROMPT-024R: "cancelling" (Dừng tất cả) is still an active batch until the worker acknowledges.
  const running = run.status === 'processing' || run.status === 'stopping' || run.status === 'cancelling'
  const locked = running || scanBusy || training || diagRunning || serverApplyRunning
  const queue = useMemo(() => {
    // PROMPT-037: a queued candidate that became skipped mid-run is no longer an actionable row.
    if (running) {
      const visibleIds = new Set(reports.map((report) => report.id))
      return run.queue.slice(run.doneCount).filter((id) => visibleIds.has(id))
    }
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
    scanBaselineGenerationRef.current = scanVersionRef.current.generation
    scanBusyRef.current = true
    setScanBusy(true)
    setScanInProgress(true)
    reportsRef.current = []
    setReports([])
    setActiveIdValue('')
    setSelectedIds([])
    clearError()
    setScanMessage('Đang quét danh sách bằng bộ quét Python…')
    try {
      const raw = await callNative<DashboardDTO>('scan_reports', configDto(), inputFileToken)
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
      scanBusyRef.current = false
      scanBaselineGenerationRef.current = null
      setScanBusy(false)
      setScanInProgress(false)
    }
  }, [acceptDashboard, clearError, configDto, inputFileToken, recordError, running])

  const chooseReportFolder = useCallback(async () => {
    clearError()
    try {
      const result = await callNative<{ cancelled: boolean; path?: string }>('choose_report_folder')
      if (!result.cancelled && result.path) setFolder(result.path)
    } catch (error) { recordError(error) }
  }, [clearError, recordError, setFolder])

  const choosePptxFile = useCallback(async () => {
    clearError()
    try {
      const result = await callNative<{ cancelled: boolean; token?: string; name?: string; folder?: string }>('choose_pptx_file')
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
      const result = await callNative<{ cancelled: boolean; path?: string }>('choose_template_file')
      if (!result.cancelled && result.path) setTemplate(result.path)
    } catch (error) { recordError(error) }
  }, [clearError, recordError, setTemplate])

  const chooseOutputFile = useCallback(async () => {
    clearError()
    try {
      const result = await callNative<{ cancelled: boolean; path?: string }>('choose_output_file')
      if (!result.cancelled && result.path) setOutput(result.path)
    } catch (error) { recordError(error) }
  }, [clearError, recordError, setOutput])

  const chooseUpdateFolder = useCallback(async () => {
    clearError()
    try {
      const result = await callNative<{ cancelled: boolean; path?: string }>('choose_update_folder')
      if (!result.cancelled && result.path) {
        updateDirty.current = true
        setUpdateState((prev) => ({ ...prev, path: result.path || '' }))
      }
    } catch (error) { recordError(error) }
  }, [clearError, recordError])

  const changeSelected = useCallback(async (restore: boolean) => {
    const ids = selectedIds.length ? selectedIds : activeId ? [activeId] : []
    if (!ids.length) return
    await runOperation(() => callNative(restore ? 'restore_reports' : 'exclude_reports', ids))
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
      await callNative('start_processing', configDto(), true)
      ollamaDirty.current = false
      updateDirty.current = false
    })
  }, [configDto, runOperation, running, scanDirty, scanned])

  const stopAfterCurrent = useCallback(() => runOperation(() => callNative('stop_after_current')), [runOperation])
  /** PROMPT-024R: "Dừng tất cả" — hủy hợp tác tại điểm an toàn gần nhất (không giết tiến trình). */
  const cancelAll = useCallback(() => runOperation(() => callNative('cancel_all')), [runOperation])
  const resetRun = clearError

  const saveConfiguration = useCallback(() => runOperation(async () => {
    await callNative('save_configuration', configDto())
    ollamaDirty.current = false
    updateDirty.current = false
  }), [configDto, runOperation])
  const checkConnection = useCallback(() => runOperation(async () => {
    await callNative('check_ollama_connection', configDto())
    ollamaDirty.current = false
    updateDirty.current = false
  }), [configDto, runOperation])
  const refreshModels = useCallback(() => runOperation(async () => {
    await callNative('refresh_ollama_models', configDto())
    ollamaDirty.current = false
    updateDirty.current = false
  }), [configDto, runOperation])
  const saveOllamaConfig = saveConfiguration
  const startDiscovery = useCallback(() => runOperation(() => callNative('start_ollama_discovery')), [runOperation])
  const stopDiscovery = useCallback(() => runOperation(() => callNative('stop_ollama_discovery')), [runOperation])
  const useServer = useCallback((host: string, port: number) => runOperation(async () => {
    await callNative('use_discovered_server', host, port)
    ollamaDirty.current = false
    setAppliedServer(`${host}:${port}`)
  }), [runOperation])

  const checkUpdate = useCallback((_startup = false) => {
    setUpdateState((current) => ({ ...current, status: 'checking', message: 'Đang kiểm tra cập nhật...',
      installEnabled: false, error: '' }))
    return runOperation(async () => {
      await callNative('check_update', configDto())
      updateDirty.current = false
      ollamaDirty.current = false
    })
  }, [callNative, configDto, runOperation])

  const requestUpdateClose = useCallback(async () => {
    if (updateCloseInFlight.current) return
    updateCloseInFlight.current = true
    try {
      const result = await callNative<unknown>('close_for_update')
      if (!isRecord(result) || result.status !== 'closing' || typeof result.message !== 'string') {
        throw new BridgeProtocolError()
      }
    } catch (error) {
      const message = safeErrorDetails(error).message
      setUpdateState((current) => ({ ...current, status: 'ready_to_restart',
        message: `Trình cập nhật đã khởi chạy nhưng cửa sổ chưa đóng. Hãy thử đóng ứng dụng để tiếp tục. ${message}`,
        error: message, installEnabled: false }))
      recordError(error)
    } finally {
      updateCloseInFlight.current = false
    }
  }, [callNative, recordError])

  useEffect(() => {
    if (update.status !== 'ready_to_restart') {
      autoUpdateCloseRequested.current = false
      return
    }
    if (autoUpdateCloseRequested.current) return
    autoUpdateCloseRequested.current = true
    void requestUpdateClose()
  }, [requestUpdateClose, update.status])

  const retryUpdateClose = useCallback(() => {
    autoUpdateCloseRequested.current = true
    void requestUpdateClose()
  }, [requestUpdateClose])

  const installUpdate = useCallback(async () => {
    if (!update.installEnabled || update.status !== 'available' || updateInstallInFlight.current) return
    updateInstallInFlight.current = true
    clearError()
    setUpdateState((current) => ({ ...current, status: 'staging', progressStage: 'PREPARING',
      message: 'Đang chuẩn bị gói cập nhật...', error: '', installEnabled: false,
      installDisabledReason: 'Đang chuẩn bị cập nhật.' }))
    try {
      const result = await callNative<unknown>('install_update')
      if (!isRecord(result) || typeof result.status !== 'string' || typeof result.message !== 'string'
        || !Object.prototype.hasOwnProperty.call(result, 'state')) {
        const message = 'Python không trả về trạng thái hợp lệ sau khi bắt đầu cập nhật.'
        setUpdateState((current) => ({ ...current, status: 'failed', message, error: message,
          installEnabled: false, installDisabledReason: message }))
        return
      }
      const next = normalizeUpdate(result.state)
      const activeStatuses: UpdateStatus[] = ['staging', 'verifying', 'installing', 'ready_to_restart']
      if (next.message === INVALID_UPDATE_MESSAGE
        || !(activeStatuses.includes(next.status) || next.status === 'failed')
        || !['staging', 'verifying', 'installing', 'ready_to_restart', 'failed'].includes(result.status)) {
        const message = 'Python không trả về trạng thái cập nhật hợp lệ sau khi bắt đầu.'
        setUpdateState((current) => ({ ...current, status: 'failed', message, error: message,
          installEnabled: false, installDisabledReason: message }))
        return
      }
      setUpdateState(next)
      if (next.status === 'failed') {
        setErrorCode('UPDATE_INSTALL_FAILED')
        setErrorMessage(next.error || next.message || result.message)
      }
    } catch (error) {
      const message = safeErrorDetails(error).message
      setUpdateState((current) => ({ ...current, status: 'failed', message, error: message,
        installEnabled: false, installDisabledReason: message }))
      recordError(error)
    } finally {
      updateInstallInFlight.current = false
    }
  }, [callNative, clearError, recordError, update.installEnabled, update.status])

  const runDiagnostics = useCallback(() => runOperation(() => callNative('run_diagnostics')), [runOperation])
  const clearLogView = useCallback(() => {
    const through = log.reduce((latest, entry) => Math.max(latest, entry.id || 0), clearLogThroughRef.current)
    clearLogThroughRef.current = through
    setLog([])
  }, [log])
  const openOutputFile = useCallback(() => runOperation(() => callNative('open_output_file')), [runOperation])
  const openOutputFolder = useCallback(() => runOperation(() => callNative('open_output_folder')), [runOperation])
  const openLogFile = useCallback(() => runOperation(() => callNative('open_log_file')), [runOperation])
  const openLogFolder = useCallback(() => runOperation(() => callNative('open_log_folder')), [runOperation])
  const openLearningFolder = useCallback(() => runOperation(() => callNative('open_learning_folder')), [runOperation])

  const setActiveManualFields = useCallback(async (id: string, fields: { vendor?: string; occurrence_date?: string }) => {
    clearError()
    try {
      const result = await callNative<{ status: string; message: string; state: DashboardDTO }>('save_manual_fields', id, fields)
      if (result.state) acceptDashboard(result.state)
      setErrorCode(result.status === 'locked' ? 'EXCEL_LOCKED' : result.status === 'error' ? 'MANUAL_FIELDS_FAILED' : '')
      setErrorMessage(result.message)
      detailStatus.current[id] = ''
    } catch (error) { recordError(error) }
  }, [acceptDashboard, clearError, recordError])
  const retryManualFields = useCallback(async (id: string) => {
    clearError()
    try {
      const result = await callNative<{ status: string; message: string; state: DashboardDTO }>('retry_manual_fields', id)
      if (result.state) acceptDashboard(result.state)
      setErrorCode(result.status === 'locked' ? 'EXCEL_LOCKED' : result.status === 'error' ? 'MANUAL_FIELDS_FAILED' : '')
      setErrorMessage(result.message)
      detailStatus.current[id] = ''
    } catch (error) { recordError(error) }
  }, [acceptDashboard, clearError, recordError])

  // PROMPT-036S: a label click updates local state IMMEDIATELY (same React tick, before any
  // bridge traffic) and persists asynchronously through the coalescing worker. The visual
  // state therefore never waits for a backend round trip, however slow the bridge is.
  const applyLocalLabel = useCallback((kind: 'image' | 'content', id: string, label: string) => {
    labelWriteEpochRef.current += 1
    const key = `${kind}:${id}`
    const existing = localLabelsRef.current.get(key)
    const candidate = (kind === 'image' ? learningStateRef.current.images : learningStateRef.current.contents)
      .find((item) => item.id === id)
    localLabelsRef.current.set(key, {
      kind, id, label,
      note: existing?.note ?? candidate?.note ?? '',
      phase: existing?.phase === 'saving' ? 'saving' : 'pending',
    })
    setLearningState((current) => {
      const list = kind === 'image' ? current.images : current.contents
      const target = list.find((item) => item.id === id)
      if (!target) return current
      const labeledDelta = target.userLabel === 'UNLABELED' && label !== 'UNLABELED' ? 1 : 0
      const counts = labeledDelta
        ? { ...current.counts, [kind]: { ...current.counts[kind], labeled: current.counts[kind].labeled + 1 } }
        : current.counts
      return kind === 'image'
        ? { ...current, images: current.images.map((item) => item.id === id ? { ...item, userLabel: label, labelPending: true } : item), counts }
        : { ...current, contents: current.contents.map((item) => item.id === id ? { ...item, userLabel: label, labelPending: true } : item), counts }
    })
    if (localLabelsRef.current.get(key)?.phase !== 'saving') persistLabel(kind, id)
  }, [persistLabel])
  const labelImage = useCallback((id: string, label: string) => applyLocalLabel('image', id, label), [applyLocalLabel])
  const labelContent = useCallback((id: string, label: string) => applyLocalLabel('content', id, label), [applyLocalLabel])
  const setNote = useCallback((kind: 'image' | 'content', id: string, note: string) => {
    const entry = localLabelsRef.current.get(`${kind}:${id}`)
    if (entry) entry.note = note
    setLearningState((current) => kind === 'image'
      ? { ...current, images: current.images.map((item) => item.id === id ? { ...item, note } : item) }
      : { ...current, contents: current.contents.map((item) => item.id === id ? { ...item, note } : item) })
  }, [])
  const saveLearningLabels = useCallback(async (kind: 'image' | 'content') => {
    const notes = kind === 'image'
      ? Object.fromEntries(learningStateRef.current.images.map((candidate) => [candidate.id, candidate.note]))
      : Object.fromEntries(learningStateRef.current.contents.map((candidate) => [candidate.id, candidate.note]))
    try {
      const result = await callNative<{ message: string; state: LearningState }>('save_learning_labels', kind, notes)
      acceptLearningState(normalizeLearning(result.state))
      setErrorCode(result.state.excelLastResult?.kind === 'locked' ? 'EXCEL_LOCKED' : '')
      setErrorMessage(result.message)
      setTraining(false)
    } catch (error) { recordError(error) }
  }, [acceptLearningState, callNative, recordError])
  const saveImageLabels = useCallback(() => saveLearningLabels('image'), [saveLearningLabels])
  const saveContentLabels = useCallback(() => saveLearningLabels('content'), [saveLearningLabels])
  // PROMPT-036S: ONE atomic bridge call (save_learning_candidate) instead of the former
  // sequential set_learning_label + save_learning_labels pair. The click is reflected
  // locally at once; on failure the choice stays visible and marked unsaved (labelPending)
  // with the error surfaced, and clicking Save again retries — it is never silently lost.
  const saveLearningCandidate = useCallback(async (kind: 'image' | 'content', id: string, label: string, note: string) => {
    const key = `${kind}:${id}`
    labelWriteEpochRef.current += 1
    localLabelsRef.current.set(key, { kind, id, label, note, phase: 'saving' })
    setLearningState((current) => kind === 'image'
      ? { ...current, images: current.images.map((item) => item.id === id ? { ...item, userLabel: label, note, labelPending: true } : item) }
      : { ...current, contents: current.contents.map((item) => item.id === id ? { ...item, userLabel: label, note, labelPending: true } : item) })
    labelWritesInFlightRef.current += 1
    try {
      // Let an in-flight fast-set for this candidate land first (backend order), then save
      // the LATEST local choice atomically. Bounded: the worker holds at most one flight.
      // Clicks during the save keep updating the entry/state immediately; the ack handler
      // below detects a superseded choice and re-persists it instead of dropping it.
      while (labelWorkersRef.current.has(key)) {
        await new Promise((resolve) => { window.setTimeout(resolve, 5) })
      }
      const latest = localLabelsRef.current.get(key)
      const sentLabel = latest?.label ?? label
      const sentNote = latest?.note ?? note
      const result = await callNative<LearningSaveAck>('save_learning_candidate', kind, id, sentLabel, sentNote)
      const current = localLabelsRef.current.get(key)
      const superseded = current && (current.label !== sentLabel || current.note !== sentNote)
      if (!superseded) {
        localLabelsRef.current.delete(key)
        setLearningState((state) => kind === 'image'
          ? { ...state, images: state.images.map((item) => item.id === id ? { ...item, labelPending: false } : item),
              excelLastResult: result.excelLastResult, excelPending: result.excelPending }
          : { ...state, contents: state.contents.map((item) => item.id === id ? { ...item, labelPending: false } : item),
              excelLastResult: result.excelLastResult, excelPending: result.excelPending })
      }
      setErrorCode(result.excelLastResult?.kind === 'locked' ? 'EXCEL_LOCKED' : '')
      setErrorMessage(result.message)
      if (superseded && current) {
        current.phase = 'pending'
        persistLabel(kind, id)
      }
      return !superseded
    } catch (error) {
      const current = localLabelsRef.current.get(key)
      if (current) {
        const superseded = current.label !== label || current.note !== note
        current.phase = superseded ? 'pending' : 'error'
        if (superseded) persistLabel(kind, id)
      }
      if (alive.current) recordError(error)
      return false
    } finally {
      labelWritesInFlightRef.current -= 1
      if (alive.current) void refreshLearning()
    }
  }, [applyLocalLabel, callNative, persistLabel, recordError, refreshLearning])
  const saveImageCandidate = useCallback((id: string, label: string, note: string) =>
    saveLearningCandidate('image', id, label, note), [saveLearningCandidate])
  const saveContentCandidate = useCallback((id: string, label: string, note: string) =>
    saveLearningCandidate('content', id, label, note), [saveLearningCandidate])
  // PROMPT-030B: one bulk save ack, not N full-state round trips. Reuse the 036S
  // local overlay/epoch/worker ordering contract, including visible failure and retry.
  const saveImageCandidates = useCallback(async (rawIds: string[], label: string): Promise<boolean> => {
    const ids = [...new Set(rawIds)]
    const candidates = ids.map((id) => learningStateRef.current.images.find((c) => c.id === id))
    const scope = (c: ImageCandidate) => `${c.reportScope || `${c.managementNumber}|${c.sourceFile}`}|${c.slide}`
    if (!ids.length || candidates.some((c) => !c)
      || new Set(candidates.map((c) => scope(c!))).size !== 1) {
      setErrorMessage('Chỉ xác nhận nhiều ảnh trong cùng báo cáo và slide.')
      return false
    }
    const keys = ids.map((id) => `image:${id}`)
    const notes = Object.fromEntries(candidates.map((c) => [c!.id, c!.note]))
    labelWriteEpochRef.current += 1
    candidates.forEach((c) => localLabelsRef.current.set(`image:${c!.id}`,
      { kind: 'image', id: c!.id, label, note: c!.note, phase: 'saving' }))
    setLearningState((state) => ({ ...state,
      images: state.images.map((c) => ids.includes(c.id) ? { ...c, userLabel: label, labelPending: true } : c),
      counts: { ...state.counts, image: { ...state.counts.image,
        labeled: state.counts.image.labeled + candidates.filter((c) => c!.userLabel === 'UNLABELED').length } },
    }))
    labelWritesInFlightRef.current += 1
    try {
      while (keys.some((key) => labelWorkersRef.current.has(key))) {
        await new Promise((resolve) => { window.setTimeout(resolve, 5) })
      }
      const result = await callNative<Omit<LearningSaveAck, 'candidateId' | 'kind'> & { candidateIds: string[] }>(
        'save_learning_images', ids, label, notes)
      const saved = new Set<string>()
      keys.forEach((key, i) => {
        const entry = localLabelsRef.current.get(key)
        if (entry?.label === label && entry.note === notes[ids[i]]) {
          localLabelsRef.current.delete(key)
          saved.add(ids[i])
        } else if (entry) {
          entry.phase = 'pending'
          persistLabel('image', ids[i])
        }
      })
      setLearningState((state) => ({ ...state,
        images: state.images.map((c) => saved.has(c.id) ? { ...c, labelPending: false } : c),
        excelLastResult: result.excelLastResult, excelPending: result.excelPending,
      }))
      setErrorCode(result.excelLastResult?.kind === 'locked' ? 'EXCEL_LOCKED' : '')
      setErrorMessage(result.message)
      return saved.size === ids.length
    } catch (error) {
      keys.forEach((key, i) => {
        const entry = localLabelsRef.current.get(key)
        if (!entry) return
        const superseded = entry.label !== label || entry.note !== notes[ids[i]]
        entry.phase = superseded ? 'pending' : 'error'
        if (superseded) persistLabel('image', ids[i])
      })
      if (alive.current) recordError(error)
      return false
    } finally {
      labelWritesInFlightRef.current -= 1
      if (alive.current) void refreshLearning()
    }
  }, [callNative, persistLabel, recordError, refreshLearning])
  const trainModels = useCallback(async () => {
    try {
      const result = await callNative<LearningState & { training?: boolean }>('train_models')
      setLearningState(normalizeLearning(result))
      setTraining(Boolean(result.training))
    } catch (error) { recordError(error) }
  }, [recordError])
  const exportLearningData = useCallback(async () => {
    clearError()
    try {
      const result = await callNative<{ cancelled?: boolean; message?: string }>('export_learning_data')
      if (result.message) setErrorMessage(result.message)
    } catch (error) { recordError(error) }
  }, [clearError, recordError])
  const applyLabelsToExcel = useCallback(async (_retry = false) => {
    try {
      const result = await callNative<{ message: string; state: LearningState }>('retry_learning_excel')
      setLearningState(normalizeLearning(result.state))
      setErrorCode(result.state.excelLastResult?.kind === 'locked' ? 'EXCEL_LOCKED' : '')
      setErrorMessage(result.message)
    } catch (error) { recordError(error) }
  }, [recordError])
  const dismissExcelResult = useCallback(async () => {
    try {
      const next = await callNative<LearningState>('dismiss_learning_excel_notice')
      setLearningState(normalizeLearning(next))
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
    force, setForce, run, reveal, running, locked, startProcessing, stopAfterCurrent, cancelAll, resetRun,
    log, logCollapsed, setLogCollapsed, clearLogView, openLogFile, openLogFolder, openOutputFile, openOutputFolder,
    ollama, setOllama, ollamaIndicator, aiStatusText, aiMode, discovering, discoveryMessage, discoveryResults,
    startDiscovery, stopDiscovery, useServer, appliedServer, serverApplyRunning, serverApplyMessage,
    checkConnection, refreshModels, saveOllamaConfig,
    update, setUpdate, checkUpdate, installUpdate, retryUpdateClose, chooseUpdateFolder,
    diagRows, diagRunning, runDiagnostics,
    learning: learningState, learningLoad, refreshLearning,
    imageReviewIndex, setImageReviewIndex, contentReviewIndex, setContentReviewIndex,
    labelImage, labelContent, setNote, saveImageLabels, saveContentLabels, saveImageCandidate, saveContentCandidate,
    saveImageCandidates,
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
