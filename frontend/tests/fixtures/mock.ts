import type {
  ContentCandidate,
  DiagnosticRow,
  DiscoveredServer,
  ImageCandidate,
  LogEntry,
  Report,
} from '../../src/types'

export const APP_VERSION = '1.3.3'
export const APP_BUILD = '016'
export const APP_TITLE = `Report Extractor — ${APP_VERSION} — Build ${APP_BUILD}`

export const DEFAULT_FOLDER = 'D:\\Reports'
export const DEFAULT_TEMPLATE = 'D:\\Reports\\Kiem_chung_mau.xlsx'
export const DEFAULT_OUTPUT = 'D:\\Reports\\Output\\Kiem_chung_Ket_qua.xlsx'

export const FOLDER_CHOICES = [
  { path: 'D:\\Reports', note: '6 báo cáo · .pptx' },
  { path: 'D:\\Reports\\2026-09', note: '6 báo cáo · .pptx' },
  { path: 'D:\\Reports\\2026-08', note: '4 báo cáo · .pptx' },
  { path: '\\\\SERVER\\ChatLuong\\Reports', note: 'thư mục mạng · LAN' },
  { path: 'E:\\CTMS\\Report', note: '3 báo cáo · .pptx' },
]

export const TEMPLATE_CHOICES = [
  { path: 'D:\\Reports\\Kiem_chung_mau.xlsx', note: 'file mẫu · sheet "Kiểm chứng"' },
  { path: 'D:\\Reports\\Kiem_chung_2026.xlsx', note: 'bản đang dùng · 129 dòng' },
  { path: '\\\\SERVER\\ChatLuong\\Kiem_chung.xlsx', note: 'thư mục mạng · chỉ đọc' },
]

export const OUTPUT_CHOICES = [
  { path: 'D:\\Reports\\Output\\Kiem_chung_Ket_qua.xlsx', note: 'chưa tồn tại · sẽ tạo mới' },
  { path: 'D:\\Reports\\Output\\Kiem_chung_T9.xlsx', note: 'đã có · 12 KB' },
  { path: 'D:\\Reports\\Output', note: 'chọn cả thư mục đầu ra' },
]

const MB = (src: string, slide: number, caption: string) => ({ src, slide, caption })

export const REPORTS: Report[] = [
  {
    id: 'r1',
    stt: 1,
    managementNumber: '260903129',
    fileName: '(CTMS)_260903129_A.pptx',
    path: 'D:\\Reports\\2026-09\\(CTMS)_260903129_A.pptx',
    vendor: '',
    manualFields: {},
    occurrenceDate: '03/09/2026',
    slides: 12,
    status: 'completed',
    shortResult: 'Nguyên nhân 2 · Cải tiến 3 · Ảnh 4',
    processedAt: '10/09/2026 14:32',
    durationText: '00:00:28',
    excelRow: 128,
    excelPath: 'D:\\Reports\\Output\\Kiem_chung_Ket_qua.xlsx',
    results: {
      qpn: MB('/mock/qpn-slide.svg', 3, 'Ảnh QPN lấy từ slide 3 của report nguồn'),
      causes: [
        'Linh kiện BOM-2291 và BOM-2292 có hình dạng ngoài giống nhau, được đặt cạnh nhau tại trạm lắp ráp số 3.',
        'Chưa có jig định vị cho thao tác ghép cụm nên công nhân lắp theo cảm tính.',
      ],
      countermeasures: [
        'Bổ sung jig định vị kim loại cho thao tác ghép cụm tại trạm lắp ráp số 3.',
        'Dán nhãn phân biệt màu cho hai mã linh kiện BOM-2291 và BOM-2292.',
        'Cập nhật hướng dẫn công việc WI-03 và đào tạo lại cho công nhân ca A, B, C.',
      ],
      temporaryRemoved: true,
      afterImages: [
        MB('/mock/after-1.jpg', 12, 'Jig định vị đã lắp tại trạm số 3'),
        MB('/mock/after-2.jpg', 12, 'Bảng hướng dẫn WI-03 đã cập nhật'),
        MB('/mock/after-3.jpg', 13, 'Vách ngăn khay linh kiện'),
        MB('/mock/after-4.jpg', 13, 'Khu vực khay linh kiện sau khi dán nhãn màu'),
      ],
    },
  },
  {
    id: 'r2',
    stt: 2,
    managementNumber: '260903030',
    fileName: '(CTMS)_260903030_B.pptx',
    path: 'D:\\Reports\\2026-09\\(CTMS)_260903030_B.pptx',
    vendor: '',
    manualFields: {},
    occurrenceDate: '03/09/2026',
    slides: 10,
    status: 'completed',
    shortResult: 'Nguyên nhân 1 · Cải tiến 2 · Ảnh 2',
    processedAt: '10/09/2026 14:34',
    durationText: '00:00:22',
    excelRow: 129,
    excelPath: 'D:\\Reports\\Output\\Kiem_chung_Ket_qua.xlsx',
    results: {
      qpn: MB('/mock/qpn-slide.svg', 2, 'Ảnh QPN lấy từ slide 2 của report nguồn'),
      causes: [
        'Hai loại linh kiện tương tự được xếp cạnh nhau, không có vách ngăn nên dễ lẫn trong quá trình lấy hàng.',
      ],
      countermeasures: [
        'Lắp vách ngăn kim loại giữa hai khay linh kiện tại trạm kiểm tra đầu vào.',
        'Dán nhãn tên và mã linh kiện trên từng khay, kèm màu phân biệt.',
      ],
      temporaryRemoved: true,
      afterImages: [
        MB('/mock/after-3.jpg', 9, 'Vách ngăn giữa hai khay'),
        MB('/mock/after-4.jpg', 10, 'Nhãn màu trên từng khay linh kiện'),
      ],
    },
  },
  {
    id: 'r3',
    stt: 3,
    managementNumber: '260903031',
    fileName: '(CTMS)_260903031_C.pptx',
    path: 'D:\\Reports\\2026-09\\(CTMS)_260903031_C.pptx',
    vendor: '',
    manualFields: {},
    occurrenceDate: '03/09/2026',
    slides: 11,
    status: 'needs_review',
    shortResult: 'Nguyên nhân 1 · Cải tiến 1 · 0 ảnh',
    warning:
      'Không tìm thấy ảnh "Sau cải tiến" trong phần CẢI TIẾN TRONG SẢN XUẤT — cần người dùng xác nhận ảnh bằng tay.',
    processedAt: '10/09/2026 14:36',
    durationText: '00:00:19',
    results: {
      qpn: MB('/mock/qpn-slide.svg', 2, 'Ảnh QPN lấy từ slide 2 của report nguồn'),
      causes: ['Chưa có điểm kiểm soát trực quan tại vị trí dễ nhầm lẫn giữa hai mã linh kiện.'],
      countermeasures: ['Bổ sung bảng kiểm tra trực quan tại trạm, do tổ trưởng xác nhận mỗi đầu ca.'],
      temporaryRemoved: true,
      afterImages: [],
    },
  },
  {
    id: 'r4',
    stt: 4,
    managementNumber: '260903032',
    fileName: '(CTMS)_260903032_D.pptx',
    path: 'D:\\Reports\\2026-09\\(CTMS)_260903032_D.pptx',
    vendor: '',
    manualFields: {},
    occurrenceDate: '03/09/2026',
    slides: 0,
    status: 'error',
    shortResult: 'Dừng ở bước đọc PPTX',
    warning:
      'Không xác định được slide QPN (đã thử nhận dạng theo từ khoá và theo mô hình). File nguồn không bị thay đổi.',
    processedAt: '10/09/2026 14:37',
    durationText: '00:00:04',
    results: { causes: [], countermeasures: [], temporaryRemoved: false, afterImages: [] },
  },
  {
    id: 'r5',
    stt: 5,
    managementNumber: '260903033',
    fileName: '(CTMS)_260903033_E.pptx',
    path: 'D:\\Reports\\2026-09\\(CTMS)_260903033_E.pptx',
    vendor: '',
    manualFields: {},
    occurrenceDate: '03/09/2026',
    slides: 12,
    status: 'skipped',
    shortResult: 'Excel đã đầy đủ — không cần cập nhật',
    processedAt: '10/09/2026 14:38',
    durationText: '00:00:02',
    results: { causes: [], countermeasures: [], temporaryRemoved: false, afterImages: [] },
  },
  {
    id: 'r6',
    stt: 6,
    managementNumber: '260826061',
    fileName: '(CTMS)_260826061_F.pptx',
    path: 'D:\\Reports\\2026-09\\(CTMS)_260826061_F.pptx',
    vendor: '',
    manualFields: {},
    occurrenceDate: '26/08/2026',
    slides: 9,
    status: 'outside_period',
    shortResult: 'Ngày phát sinh 26/08/2026 ngoài khoảng đã chọn',
    warning: 'Báo cáo nằm ngoài khoảng thời gian đang chọn nên chưa mở PPTX.',
    results: { causes: [], countermeasures: [], temporaryRemoved: false, afterImages: [] },
  },
  {
    id: 'r7',
    stt: 7,
    managementNumber: '260904001',
    fileName: '(CTMS)_260904001_G.pptx',
    path: 'D:\\Reports\\2026-09\\(CTMS)_260904001_G.pptx',
    vendor: '',
    manualFields: {},
    occurrenceDate: '04/09/2026',
    slides: 13,
    status: 'new_row',
    shortResult: 'Sẽ thêm dòng mới vào file kết quả',
    results: {
      qpn: MB('/mock/qpn-slide.svg', 3, 'Ảnh QPN lấy từ slide 3 của report nguồn'),
      causes: ['Hai mã linh kiện gần giống nhau được bày trong cùng một khay tại trạm số 3.'],
      countermeasures: ['Thêm vách ngăn và nhãn màu cho từng khay linh kiện.'],
      temporaryRemoved: true,
      afterImages: [
        MB('/mock/after-3.jpg', 14, 'Vách ngăn khay linh kiện'),
        MB('/mock/after-4.jpg', 14, 'Nhãn màu trên từng khay'),
      ],
    },
  },
  {
    id: 'r8',
    stt: 8,
    managementNumber: '260904002',
    fileName: '(CTMS)_260904002_H.pptx',
    path: 'D:\\Reports\\2026-09\\(CTMS)_260904002_H.pptx',
    vendor: '',
    manualFields: {},
    occurrenceDate: '04/09/2026',
    slides: 9,
    status: 'waiting',
    shortResult: 'Trong hàng đợi',
    results: {
      qpn: MB('/mock/qpn-slide.svg', 2, 'Ảnh QPN lấy từ slide 2 của report nguồn'),
      causes: [],
      countermeasures: [],
      temporaryRemoved: false,
      afterImages: [],
    },
  },
  {
    id: 'r9',
    stt: 9,
    managementNumber: '260904003',
    fileName: '(CTMS)_260904003_I.pptx',
    path: 'D:\\Reports\\2026-09\\(CTMS)_260904003_I.pptx',
    vendor: '',
    manualFields: {},
    occurrenceDate: '04/09/2026',
    slides: 8,
    status: 'excluded',
    shortResult: 'Loại khỏi danh sách xử lý — không xóa file nguồn',
    results: { causes: [], countermeasures: [], temporaryRemoved: false, afterImages: [] },
  },
  {
    id: 'r10',
    stt: 10,
    managementNumber: '260904004',
    fileName: '(CTMS)_260904004_K.pptx',
    path: 'D:\\Reports\\2026-09\\(CTMS)_260904004_K.pptx',
    vendor: '',
    manualFields: {},
    occurrenceDate: '04/09/2026',
    slides: 11,
    status: 'fast_skip',
    shortResult: 'Đã xử lý gần đây (12/09/2026 09:10)',
    processedAt: '12/09/2026 09:10',
    durationText: '00:00:01',
    results: { causes: [], countermeasures: [], temporaryRemoved: false, afterImages: [] },
  },
]

/** Kết quả mô phỏng sau khi chạy xử lý một báo cáo. */
export const RUN_OUTCOME: Record<string, { status: Report['status']; shortResult: string; durationText: string }> = {
  r7: { status: 'completed', shortResult: 'Nguyên nhân 1 · Cải tiến 1 · Ảnh 2', durationText: '00:00:26' },
  r8: {
    status: 'needs_review',
    shortResult: 'Nguyên nhân 2 · Cải tiến 1 · 0 ảnh',
    durationText: '00:00:31',
  },
}

export const INITIAL_LOG: LogEntry[] = [
  { time: '14:32:01', level: 'INFO', text: 'Bắt đầu xử lý file: (CTMS)_260903129_A.pptx' },
  { time: '14:32:02', level: 'INFO', text: 'Phân tích cấu trúc PPTX …' },
  { time: '14:32:05', level: 'INFO', text: 'Trích xuất nguyên nhân … (2 mục)' },
  { time: '14:32:10', level: 'INFO', text: 'Trích xuất cải tiến … (3 mục)' },
  { time: '14:32:16', level: 'INFO', text: 'Trích xuất hình ảnh … (4 ảnh)' },
  { time: '14:32:24', level: 'INFO', text: 'Ghi kết quả Excel …' },
  { time: '14:32:28', level: 'INFO', text: 'Hoàn thành: (CTMS)_260903129_A.pptx (00:00:28)' },
  { time: '14:34:02', level: 'INFO', text: 'Bắt đầu xử lý file: (CTMS)_260903030_B.pptx' },
  { time: '14:34:24', level: 'INFO', text: 'Hoàn thành: (CTMS)_260903030_B.pptx (00:00:22)' },
  {
    time: '14:36:11',
    level: 'WARN',
    text: 'Cần kiểm tra: (CTMS)_260903031_C.pptx — không tìm thấy ảnh "Sau cải tiến" trong phần CẢI TIẾN TRONG SẢN XUẤT',
  },
  { time: '14:37:20', level: 'ERROR', text: 'Lỗi: (CTMS)_260903032_D.pptx — không xác định được slide QPN' },
  { time: '14:38:04', level: 'INFO', text: 'Bỏ qua: (CTMS)_260903033_E.pptx — Excel đã đầy đủ' },
  { time: '14:38:12', level: 'INFO', text: 'Đã dừng theo yêu cầu. Hàng đợi còn 3 báo cáo chưa xử lý.' },
]

export const LOG_LEVEL_STYLE: Record<LogEntry['level'], string> = {
  INFO: 'text-[#1d63b8]',
  WARN: 'text-[#b45309]',
  ERROR: 'text-[#b91c1c]',
  DEBUG: 'text-[#5b7290]',
}

export const FALLBACK_NOTICE =
  'Ollama không khả dụng → chương trình dùng phân tích từ khoá (heuristic) và vẫn ghi kết quả bình thường bằng cách sao chép văn bản, ảnh gốc từ report.'

export const MODEL_CHOICES = ['qwen3:4b', 'qwen3:8b', 'qwen2.5:7b', 'llama3.1:8b']

export const DISCOVERED_SERVERS: DiscoveredServer[] = [
  { host: '192.168.1.24', port: 11434, models: 6, version: '0.12.3', note: 'Đã xác minh /api/tags' },
  { host: '192.168.1.51', port: 11434, models: 4, version: '0.11.7', note: 'Đã xác minh /api/tags' },
  { host: '10.0.0.18', port: 11434, models: 9, version: '0.12.1', note: 'Đã xác minh /api/tags' },
]

export const UPDATE_PATH_DEFAULT = 'D:\\ReportExtractor_Update'

/** Kết quả bốn mục kiểm tra bắt buộc của tab Cài đặt → Chẩn đoán. */
export const DIAGNOSTIC_IDLE: DiagnosticRow[] = [
  { label: 'Ollama', value: 'Chưa chạy kiểm tra', state: 'idle' },
  { label: 'Template Excel', value: 'Chưa chạy kiểm tra', state: 'idle' },
  { label: 'Thư mục đầu ra', value: 'Chưa chạy kiểm tra', state: 'idle' },
  { label: 'Bộ dựng ảnh QPN', value: 'Chưa chạy kiểm tra', state: 'idle' },
]

export const DIAGNOSTIC_RESULT: DiagnosticRow[] = [
  { label: 'Ollama', value: 'Đã kết nối — http://192.168.1.24:11434 (6 model, qwen3:4b khả dụng)', state: 'ok' },
  { label: 'Template Excel', value: 'Hợp lệ — sheet "Kiểm chứng", dòng tiêu đề 2, cột A–N', state: 'ok' },
  { label: 'Thư mục đầu ra', value: 'Ghi được — D:\\Reports\\Output', state: 'ok' },
  {
    label: 'Bộ dựng ảnh QPN',
    value: 'PowerPoint COM: có · LibreOffice: không · Bộ dựng tích hợp: sẵn sàng',
    state: 'warn',
  },
]

/* ------------------------------------------------------- Học cải tiến: ứng viên ảnh */

export const IMAGE_CANDIDATES: ImageCandidate[] = [
  {
    id: 'img1',
    sourceFile: '(CTMS)_260903129_A.pptx',
    managementNumber: '260903129',
    slide: 12,
    pictureId: '4',
    src: '/mock/after-1.jpg',
    bounds: { x: 6, y: 34, w: 40, h: 38 },
    decision: 'Sau cải tiến',
    confidenceBand: 'high',
    confidence: 0.86,
    evidence: [
      'Ảnh nằm trong mục CẢI TIẾN TRONG SẢN XUẤT của slide 12',
      'Có mũi tên chỉ từ ảnh "Trước cải tiến" sang ảnh này',
      'Chữ gần ảnh: "Sau cải tiến"',
    ],
    excelEligible: true,
    nearbyText: 'Sau cải tiến — jig định vị trạm số 3',
    userLabel: 'UNLABELED',
    note: '',
  },
  {
    id: 'img2',
    sourceFile: '(CTMS)_260903129_A.pptx',
    managementNumber: '260903129',
    slide: 12,
    pictureId: '5',
    src: '/mock/after-2.jpg',
    bounds: { x: 52, y: 34, w: 42, h: 38 },
    decision: 'Sau cải tiến',
    confidenceBand: 'high',
    confidence: 0.81,
    evidence: ['Nằm cùng nhóm ảnh sau cải tiến', 'Có tiêu đề mục cải tiến phía trên'],
    excelEligible: true,
    nearbyText: 'Sau cải tiến — bảng hướng dẫn WI-03',
    userLabel: 'UNLABELED',
    note: '',
  },
  {
    id: 'img3',
    sourceFile: '(CTMS)_260903129_A.pptx',
    managementNumber: '260903129',
    slide: 13,
    pictureId: '6',
    src: '/mock/after-3.jpg',
    bounds: { x: 54, y: 30, w: 40, h: 42 },
    decision: 'Cần kiểm tra',
    confidenceBand: 'medium',
    confidence: 0.58,
    evidence: [
      'Ảnh nằm ở mép phải slide, chưa xác định được mục sở hữu',
      'Không thấy chữ "Sau cải tiến" trong cùng khối',
    ],
    excelEligible: false,
    eligibilityReason: 'cần xác nhận độc lập',
    nearbyText: 'Vách ngăn khay',
    userLabel: 'UNLABELED',
    note: '',
  },
  {
    id: 'img4',
    sourceFile: '(CTMS)_260903030_B.pptx',
    managementNumber: '260903030',
    slide: 9,
    pictureId: '3',
    src: '/mock/after-4.jpg',
    bounds: { x: 8, y: 36, w: 38, h: 36 },
    decision: 'Sau cải tiến',
    confidenceBand: 'high',
    confidence: 0.79,
    evidence: ['Thuộc phần cải tiến trong sản xuất', 'Chữ gần ảnh: "Sau khi bổ sung nhãn màu"'],
    excelEligible: true,
    nearbyText: 'Nhãn màu trên từng khay',
    userLabel: 'UNLABELED',
    note: '',
  },
  {
    id: 'img5',
    sourceFile: '(CTMS)_260903030_B.pptx',
    managementNumber: '260903030',
    slide: 10,
    pictureId: '2',
    src: '/mock/before-1.jpg',
    bounds: { x: 10, y: 32, w: 40, h: 40 },
    decision: 'Trước cải tiến',
    confidenceBand: 'high',
    confidence: 0.9,
    evidence: ['Có chữ "Hiện trạng" phía trên ảnh', 'Nằm trước ảnh sau cải tiến theo thứ tự đọc'],
    excelEligible: false,
    eligibilityReason: 'chỉ lấy ảnh "Sau cải tiến" của phần cải tiến sản xuất',
    nearbyText: 'Hiện trạng trước cải tiến',
    userLabel: 'UNLABELED',
    note: '',
  },
  {
    id: 'img6',
    sourceFile: '(CTMS)_260903031_C.pptx',
    managementNumber: '260903031',
    slide: 11,
    pictureId: '1',
    src: '/mock/before-2.jpg',
    bounds: { x: 12, y: 30, w: 44, h: 42 },
    decision: 'Cần kiểm tra',
    confidenceBand: 'low',
    confidence: 0.44,
    evidence: [
      'Không tìm thấy văn bản "Sau cải tiến" gần ảnh',
      'Ảnh nằm trong khối có tiêu đề mục cải tiến nhưng không có mũi tên chỉ dẫn',
    ],
    excelEligible: false,
    eligibilityReason: 'cần xác nhận độc lập',
    nearbyText: 'Bảng kiểm tra trực quan',
    userLabel: 'UNLABELED',
    note: '',
  },
]

/* ------------------------------------------------- Học cải tiến: ứng viên nội dung */

export const CONTENT_CANDIDATES: ContentCandidate[] = [
  {
    id: 'ct1',
    sourceFile: '(CTMS)_260903129_A.pptx',
    managementNumber: '260903129',
    slide: 12,
    blockId: '18',
    text:
      'Bổ sung jig định vị kim loại cho thao tác ghép cụm tại trạm lắp ráp số 3.\n' +
      'Dán nhãn phân biệt màu cho hai mã linh kiện BOM-2291 và BOM-2292.\n' +
      'Cập nhật hướng dẫn công việc WI-03 và đào tạo lại cho công nhân ca A, B, C.',
    decision: 'Nội dung cải tiến',
    confidenceBand: 'high',
    confidence: 0.88,
    evidence: ['Khối chữ nằm trong mục CẢI TIẾN TRONG SẢN XUẤT', 'Không thuộc phần XỬ LÝ TẠM THỜI'],
    section: 'CẢI TIẾN TRONG SẢN XUẤT',
    nearestTitle: 'Đối sách cải tiến',
    userLabel: 'UNLABELED',
    note: '',
  },
  {
    id: 'ct2',
    sourceFile: '(CTMS)_260903129_A.pptx',
    managementNumber: '260903129',
    slide: 12,
    blockId: '21',
    text:
      'Xử lý tạm thời: phân loại lại toàn bộ linh kiện trong ngày và kiểm tra 100% bằng mắt trước khi chuyển trạm.',
    decision: 'Không lấy',
    confidenceBand: 'high',
    confidence: 0.92,
    evidence: ['Khối chữ bắt đầu bằng nhãn "Xử lý tạm thời"', 'Thuộc phần XỬ LÝ TẠM THỜI — loại khỏi nội dung đối sách'],
    section: 'XỬ LÝ TẠM THỜI',
    nearestTitle: 'Xử lý tạm thời',
    userLabel: 'UNLABELED',
    note: '',
  },
  {
    id: 'ct3',
    sourceFile: '(CTMS)_260903030_B.pptx',
    managementNumber: '260903030',
    slide: 10,
    blockId: '14',
    text:
      'Lắp vách ngăn kim loại giữa hai khay linh kiện tại trạm kiểm tra đầu vào.\n' +
      'Dán nhãn tên và mã linh kiện trên từng khay, kèm màu phân biệt.',
    decision: 'Nội dung cải tiến',
    confidenceBand: 'high',
    confidence: 0.84,
    evidence: ['Khối chữ nằm trong mục CẢI TIẾN TRONG SẢN XUẤT'],
    section: 'CẢI TIẾN TRONG SẢN XUẤT',
    nearestTitle: 'Đối sách cải tiến',
    userLabel: 'UNLABELED',
    note: '',
  },
  {
    id: 'ct4',
    sourceFile: '(CTMS)_260903030_B.pptx',
    managementNumber: '260903030',
    slide: 8,
    blockId: '6',
    text: 'Tiêu chuẩn kiểm tra linh kiện đầu vào: kiểm tra 5 mẫu mỗi lô, ghi nhật ký kiểm tra tại trạm.',
    decision: 'Cần kiểm tra',
    confidenceBand: 'medium',
    confidence: 0.61,
    evidence: ['Khối chữ thuộc phần tiêu chuẩn / SOP', 'Có thể là nội dung cải tiến nếu người dùng xác nhận'],
    section: 'tiêu chuẩn / SOP',
    nearestTitle: 'Tiêu chuẩn kiểm tra',
    userLabel: 'UNLABELED',
    note: '',
  },
  {
    id: 'ct5',
    sourceFile: '(CTMS)_260903031_C.pptx',
    managementNumber: '260903031',
    slide: 11,
    blockId: '9',
    text:
      'Bổ sung bảng kiểm tra trực quan tại trạm, tổ trưởng xác nhận và ký trước mỗi đầu ca sản xuất.',
    decision: 'Cần kiểm tra',
    confidenceBand: 'low',
    confidence: 0.47,
    evidence: [
      'Khối chữ nằm sát biên phải slide, chưa xác định được mục sở hữu',
      'Không có tiêu đề mục cải tiến phía trên trong cùng khối',
    ],
    section: 'chưa xác định',
    nearestTitle: '—',
    userLabel: 'UNLABELED',
    note: '',
  },
]

export const LEARNING_COUNTS = {
  image: { total: 24, labeled: 18 },
  content: { total: 12, labeled: 7 },
  /** Số mẫu tối thiểu để huấn luyện một mô hình. */
  minSamples: 20,
}

export const LEARNING_MODEL_STATUS = {
  image: 'Chưa huấn luyện — chưa đủ dữ liệu học (18/20 nhãn cần thiết). Quy tắc hiện tại vẫn được dùng.',
  content: 'Đã huấn luyện — 7 mẫu, cập nhật 10/09/2026 14:05. Dùng để xếp hạng vùng nội dung, không viết lại nội dung.',
}
