/** JSON-safe view models returned by the Python application service. */

export type TabKey = 'reports' | 'settings' | 'learning'

/** Giai đoạn xử lý một báo cáo (STAGE_LABELS_VI). */
export type StageKey =
  | 'waiting'
  | 'reading'
  | 'analyzing'
  | 'analyzing_heuristic'
  | 'extracting'
  | 'extracting_qpn'
  | 'extracting_images'
  | 'writing_excel'

export const STAGE_LABELS: Record<StageKey, string> = {
  waiting: 'Đang chờ',
  reading: 'Đang đọc PPTX',
  analyzing: 'Đang phân tích Qwen',
  analyzing_heuristic: 'Đang phân tích (từ khoá, không AI)',
  extracting: 'Đang trích xuất nguyên nhân / đối sách cải tiến',
  extracting_qpn: 'Đang trích xuất QPN',
  extracting_images: 'Đang trích xuất hình ảnh cải tiến',
  writing_excel: 'Đang ghi Excel',
}

/** Thứ tự giai đoạn thật của pipeline (không nội suy). */
export const STAGE_ORDER: StageKey[] = [
  'reading',
  'analyzing',
  'extracting',
  'extracting_qpn',
  'extracting_images',
  'writing_excel',
]

/** Tỉ lệ hoàn thành của MỘT báo cáo khi đạt tới giai đoạn thật (STAGE_WEIGHTS). */
export const STAGE_WEIGHTS: Record<string, number> = {
  reading: 0.05,
  analyzing: 0.15,
  analyzing_heuristic: 0.15,
  extracting: 0.45,
  extracting_qpn: 0.55,
  extracting_images: 0.7,
  writing_excel: 0.85,
}

/** Trạng thái cuối của một báo cáo trong danh sách (STATUS_VI). */
export type StatusKey =
  | 'waiting'
  | 'processing'
  | 'completed'
  | 'needs_review'
  | 'error'
  | 'skipped'
  | 'cancelled'
  | 'outside_period'
  | 'source_duplicate'
  | 'fast_skip'
  | 'excluded'
  | 'new_row'

export const STATUS_LABEL: Record<StatusKey, string> = {
  waiting: 'Đang chờ',
  processing: 'Đang xử lý',
  completed: 'Hoàn thành',
  needs_review: 'Cần kiểm tra',
  error: 'Lỗi',
  skipped: 'Bỏ qua — đã cập nhật',
  /** PROMPT-024R: người dùng chọn "Dừng tất cả" — không phải lỗi xử lý. */
  cancelled: 'Đã hủy',
  outside_period: 'Bỏ qua ngoài thời gian xử lý',
  source_duplicate: 'Trùng Management Number trong folder',
  fast_skip: 'Bỏ qua nhanh — đã xử lý gần đây',
  excluded: 'Đã loại thủ công',
  new_row: 'Sẽ xử lý — Management Number mới',
}

/** Nhóm thống kê bắt buộc: Hoàn thành · Cần kiểm tra · Lỗi · Bỏ qua. */
export type StatBucket = 'completed' | 'needs_review' | 'error' | 'skipped'

export const BUCKET_OF: Record<StatusKey, StatBucket | 'other'> = {
  completed: 'completed',
  processing: 'other',
  new_row: 'other',
  waiting: 'other',
  needs_review: 'needs_review',
  error: 'error',
  // PROMPT-024R: cancellation is a user choice — never counted in the Lỗi (failure) bucket.
  cancelled: 'other',
  skipped: 'skipped',
  outside_period: 'skipped',
  source_duplicate: 'skipped',
  fast_skip: 'skipped',
  excluded: 'other',
}

/**
 * PROMPT-037 / PROMPT-037R: the ONE authoritative rule for rows shown in “Danh sách báo cáo”.
 * Every status whose UI meaning is “Bỏ qua” is hidden from the list, its totals, selection and
 * the active row: `skipped` (“Bỏ qua — đã cập nhật”), `outside_period` (“Bỏ qua ngoài thời gian
 * xử lý”), `source_duplicate` and `fast_skip` (“Bỏ qua nhanh — đã xử lý gần đây”). The DTO still
 * carries them: the backend keeps the classification for processing, duplicate handling, scan
 * decisions and incremental transitions. Presentation-only; nothing is blacklisted by path, so a
 * later scan that makes a report visible/actionable shows it again.
 */
export const HIDDEN_REPORT_STATUSES: ReadonlySet<StatusKey> = new Set<StatusKey>([
  'skipped',
  'outside_period',
  'source_duplicate',
  'fast_skip',
])

/**
 * PROMPT-038: canonical form of a DTO status for the visibility rule. The checked-in backend
 * already emits exact keys (`outside_period`, …), but the single authoritative predicate must
 * not silently leak a “Bỏ qua” row if a future DTO pathway emits different casing/whitespace.
 */
export function normalizeReportStatus(status: unknown): string {
  return typeof status === 'string' ? status.trim().toLowerCase() : ''
}

export function isVisibleReport(report: Pick<Report, 'status'>): boolean {
  return !HIDDEN_REPORT_STATUSES.has(normalizeReportStatus(report.status) as StatusKey)
}

/** O(N): visible rows in incoming order, STT renumbered over the visible list only. */
export function presentReports(list: readonly Report[]): Report[] {
  const out: Report[] = []
  for (const report of list) {
    if (!isVisibleReport(report)) continue
    out.push(report.stt === out.length + 1 ? report : { ...report, stt: out.length + 1 })
  }
  return out
}

export type QpnImage = {
  src: string
  slide: number
  caption: string
}

export type AfterImage = {
  src: string
  slide: number
  caption: string
}

export type Report = {
  id: string
  stt: number
  managementNumber: string
  fileName: string
  path: string
  /** Ngày phát sinh hiển thị dd/mm/yyyy; có thể là ngày đã được người dùng sửa. */
  occurrenceDate: string
  vendor: string
  manualFields: { vendor?: string; occurrence_date?: string }
  manualPending?: boolean
  slides: number
  status: StatusKey
  /** Cột "Cảnh báo / kết quả" ngắn gọn trong bảng. */
  shortResult: string
  /** Cảnh báo cụ thể hiển thị trong khung chi tiết. */
  warning?: string
  processedAt?: string
  durationText?: string
  excelRow?: number
  excelPath?: string
  results: {
    qpn?: QpnImage
    /** Nguyên nhân — giữ nguyên văn. */
    causes: string[]
    /** Nội dung đối sách cải tiến — đã loại phần "Xử lý tạm thời". */
    countermeasures: string[]
    temporaryRemoved: boolean
    /** Chỉ ảnh "Sau cải tiến" thuộc phần CẢI TIẾN TRONG SẢN XUẤT. */
    afterImages: AfterImage[]
  }
}

export type LogEntry = {
  id?: number
  time: string
  level: 'INFO' | 'WARN' | 'ERROR' | 'DEBUG'
  text: string
}

export type OllamaConnection = {
  server: string
  port: string
  model: string
  checked: 'unchecked' | 'ok' | 'fail'
  models: string[]
  message: string
}

export type DiscoveredServer = {
  host: string
  port: number
  models: number
  version: string
  note: string
}

export type DiagnosticRow = {
  label: string
  value: string
  state: 'ok' | 'warn' | 'error' | 'idle'
}

/* ---------------------------------------------------------------- Excel mặc định (PROMPT-040) */

export type ExcelKind = 'template' | 'output'

/** `unconfigured` = chưa từng chọn; `checking` = đang kiểm tra hợp lệ; các trạng thái còn lại là lỗi thật. */
export type ExcelState = 'unconfigured' | 'checking' | 'ok' | 'missing' | 'inaccessible' | 'invalid'

/** Một file Excel đã được chọn một lần và ghi vào config người dùng (không dùng localStorage). */
export type ExcelKindState = {
  kind: ExcelKind
  label: string
  path: string
  name: string
  folder: string
  state: ExcelState
  message: string
  /** false = giá trị trong config.json chưa khớp lựa chọn hiện tại (ví dụ thư mục config chỉ đọc). */
  persisted: boolean
}

export type ExcelSelectionState = { template: ExcelKindState; output: ExcelKindState }

export const EXCEL_STATE_LABEL: Record<ExcelState, string> = {
  unconfigured: 'Chưa chọn',
  checking: 'Đang kiểm tra',
  ok: 'Dùng được',
  missing: 'Không tìm thấy',
  inaccessible: 'Không truy cập được',
  invalid: 'Không hợp lệ',
}

/* ---------------------------------------------------------------- Học cải tiến */

export type LabelValue = 'UNLABELED' | string

/** Ảnh cải tiến: nhãn xác nhận (LABEL_VI của app/image_learning.py). */
export const IMAGE_LABELS: { key: string; label: string }[] = [
  { key: 'AFTER', label: 'Sau cải tiến' },
  { key: 'BEFORE', label: 'Trước cải tiến' },
  { key: 'CONTROL', label: 'Kiểm tra / Kiểm soát' },
  { key: 'IGNORE', label: 'Không lấy' },
]

export const CONTENT_LABELS: { key: string; label: string }[] = [
  { key: 'IMPROVEMENT_CONTENT', label: 'Nội dung cải tiến' },
  { key: 'EXCLUDE_CONTENT', label: 'Không lấy' },
]

export const CONFIDENCE_LABEL: Record<'high' | 'medium' | 'low', string> = {
  high: 'Cao',
  medium: 'Trung bình',
  low: 'Thấp',
}

export type ImageCandidate = {
  /** Scoped evidence identity; canonical `id` remains backward-compatible in the Python DTO. */
  reviewId?: string
  reportScope?: string
  persisted?: boolean
  id: string
  sourceFile: string
  managementNumber: string
  slide: number
  pictureId: string
  src: string
  /** Khung ảnh trên slide, tính theo % kích thước slide. */
  bounds: { x: number; y: number; w: number; h: number }
  /** PROMPT-024R: kích thước slide gốc (EMU) + khung mục tiêu EMU để map toạ độ chính xác. */
  slideWidth?: number
  slideHeight?: number
  targetBbox?: { x: number; y: number; width: number; height: number }
  targetBboxPct?: { x: number; y: number; w: number; h: number }
  targetKind?: string
  /** PROMPT-024R: ảnh TOÀN slide đã render (data URI) — nền ngữ cảnh cho vị trí thật của mục tiêu. */
  slidePreview?: string
  slidePreviewWidth?: number
  slidePreviewHeight?: number
  /**
   * PROMPT-027 §4/§6: backend ĐÃ thật sự render ảnh xem trước ("powerpoint" | "libreoffice" | "builtin").
   * Chỉ PowerPoint là trung thực từng điểm ảnh với slide gốc; backend dự phòng phải được báo rõ để UI
   * hiện cảnh báo giảm độ trung thực thay vì giả vờ là kết quả PowerPoint. Không chứa đường dẫn file.
   */
  slidePreviewBackend?: string
  slidePreviewFaithful?: boolean
  /**
   * PROMPT-027 §13/§14/§15: VÙNG XUẤT EXCEL — vùng "Sau cải tiến" theo ImprovementItem mà pipeline
   * sản xuất SẼ crop đưa vào Excel nếu xử lý file này ngay bây giờ. Do Python tính từ
   * ImprovementVisualRegion.bbox; React KHÔNG BAO GIỜ tự suy ra từ bounds của ảnh ứng viên.
   * Một vùng có thể chứa nhiều ảnh After + caption + chú thích; null khi ảnh này không thuộc vùng nào.
   */
  evidenceRegionBbox?: { x: number; y: number; width: number; height: number } | null
  evidenceRegionBboxPct?: { x: number; y: number; w: number; h: number } | null
  evidenceRegionKind?: string
  evidenceRegionItemId?: string
  evidenceRegionItemIndex?: number
  evidenceRegionItemHeading?: string
  evidenceRegionPictureCount?: number
  /** PROMPT-027R §24: shape ids of the pictures that form the production region — display/debug only; React never
   *  derives any geometry from them. */
  evidenceRegionPictureIds?: number[]
  evidenceRegionId?: string
  evidenceRegionBlockIndex?: number
  evidenceRegionBoundarySource?: string
  evidenceRegionNextHeadingTop?: number | null
  /** PROMPT-025: định danh mục cải tiến (một slide có thể có nhiều mục). */
  itemId?: string
  itemIndex?: number
  itemHeading?: string
  decision: string
  confidenceBand: 'high' | 'medium' | 'low'
  confidence: number
  evidence: string[]
  excelEligible: boolean
  eligibilityReason?: string
  nearbyText?: string
  userLabel: LabelValue
  labelPending?: boolean
  note: string
}

export type ContentCandidate = {
  id: string
  sourceFile: string
  managementNumber: string
  slide: number
  blockId: string
  /** Nguyên văn nội dung khối chữ — không tóm tắt, không viết lại. */
  text: string
  decision: string
  confidenceBand: 'high' | 'medium' | 'low'
  confidence: number
  evidence: string[]
  section: string
  nearestTitle: string
  userLabel: LabelValue
  labelPending?: boolean
  note: string
}

export type ProcessingRun = {
  /** PROMPT-024R: "cancelling" = "Dừng tất cả" đã yêu cầu, chờ worker xác nhận ở điểm an toàn. */
  status: 'idle' | 'processing' | 'stopping' | 'cancelling' | 'done'
  queue: string[]
  /** Chỉ số báo cáo đang xử lý trong hàng đợi. */
  index: number
  /** Số báo cáo đã xong trong lần chạy này. */
  doneCount: number
  stage: StageKey
  percent: number
  currentFile: string
  elapsedSec: number
  remainSec: number | null
  startedAt: number | null
  finishedAt: number | null
  /** true khi ít nhất một báo cáo đã kết thúc trong phiên này (để hiển thị ETA). */
  hasSamples: boolean
  stopped?: boolean
  /** PROMPT-024R: lượt chạy này có yêu cầu "Dừng tất cả" hay không + số báo cáo đã hủy. */
  cancelRequested?: boolean
  cancelledCount?: number
}
