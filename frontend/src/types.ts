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
  skipped: 'skipped',
  outside_period: 'skipped',
  source_duplicate: 'skipped',
  fast_skip: 'skipped',
  excluded: 'other',
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
  id: string
  sourceFile: string
  managementNumber: string
  slide: number
  pictureId: string
  src: string
  /** Khung ảnh trên slide, tính theo % kích thước slide. */
  bounds: { x: number; y: number; w: number; h: number }
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
  status: 'idle' | 'processing' | 'stopping' | 'done'
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
}
