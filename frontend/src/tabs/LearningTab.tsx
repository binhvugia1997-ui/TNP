import { useEffect, useMemo, useRef, useState } from 'react'
import {
  AlertTriangle,
  BarChart3,
  CheckCircle2,
  ChevronLeft,
  ChevronRight,
  Cpu,
  Database,
  FileSpreadsheet,
  FolderOpen,
  Images,
  Info,
  Loader2,
  Maximize2,
  Minus,
  Plus,
  RefreshCw,
  Save,
  ShieldAlert,
  Text,
  Upload,
} from 'lucide-react'
import { Button, Card, KeyValue, ProgressBar, TextInput, ToneChip } from '../components/ui'
import { useStore, type LearningState } from '../state/store'
import { CONFIDENCE_LABEL, IMAGE_LABELS, type ImageCandidate } from '../types'
import { cn } from '../lib/utils'

type SectionKey = 'overview' | 'imageReview' | 'contentReview' | 'models'

const MENU: { key: SectionKey; label: string; icon: typeof BarChart3 }[] = [
  { key: 'overview', label: 'Tổng quan', icon: BarChart3 },
  { key: 'imageReview', label: 'Kiểm tra ảnh cải tiến', icon: Images },
  { key: 'contentReview', label: 'Kiểm tra nội dung cải tiến', icon: Text },
  { key: 'models', label: 'Mô hình & dữ liệu', icon: Database },
]

export function LearningTab() {
  const s = useStore()
  const [section, setSection] = useState<SectionKey>('overview')

  return (
    <div className="flex min-h-full flex-col gap-3 p-3 lg:flex-row">
      <aside className="w-full shrink-0 lg:w-[224px]">
        <Card title="Học cải tiến" dense bodyClassName="p-1.5">
          <nav className="flex flex-col gap-0.5">
            {MENU.map(({ key, label, icon: Icon }) => (
              <button
                key={key}
                type="button"
                onClick={() => setSection(key)}
                className={cn(
                  'flex items-center gap-2 rounded-sm2 px-2 py-1.5 text-left text-base2',
                  section === key ? 'bg-brand-50 font-semibold text-brand-700' : 'text-body hover:bg-[#f4f7fb]',
                )}
              >
                <Icon className={cn('h-4 w-4 shrink-0', section === key ? 'text-brand-600' : 'text-muted')} />
                {label}
              </button>
            ))}
          </nav>
          <p className="px-2 py-1.5 text-xxs leading-[15px] text-muted">
            Học từ nhãn xác nhận của người dùng về ảnh và vùng nội dung trích xuất.
          </p>
        </Card>
      </aside>

      <div className="flex min-w-0 flex-1 flex-col gap-3">
        {section === 'overview' && <Overview learning={s.learning} onGo={setSection} />}
        {section === 'imageReview' && <ImageReview />}
        {section === 'contentReview' && <ContentReview />}
        {section === 'models' && <ModelAndData />}
      </div>
    </div>
  )
}

/* ------------------------------------------------------------------ Tổng quan */

function Overview({ learning, onGo }: { learning: LearningState; onGo: (s: SectionKey) => void }) {
  const imgPending = learning.images.filter((c) => c.userLabel === 'UNLABELED').length
  const contentPending = learning.contents.filter((c) => c.userLabel === 'UNLABELED').length
  const imgLabeled = learning.counts.image.labeled
  const contentLabeled = learning.counts.content.labeled
  const imageModel = modelSummary(learning.modelStatus.image)
  const contentModel = modelSummary(learning.modelStatus.content)

  return (
    <>
      <div className="grid grid-cols-1 gap-3 xl:grid-cols-2">
        <LearningGroup
          title="Dữ liệu học ảnh"
          icon={<Images className="h-4 w-4" />}
          total={learning.counts.image.total}
          labeled={imgLabeled}
          pending={imgPending}
          note="Ảnh ứng viên trích từ report; nhãn xác nhận quyết định ảnh nào được đưa vào Excel."
          action={<Button onClick={() => onGo('imageReview')}>Kiểm tra ảnh cải tiến</Button>}
        />
        <LearningGroup
          title="Dữ liệu học nội dung"
          icon={<Text className="h-4 w-4" />}
          total={learning.counts.content.total}
          labeled={contentLabeled}
          pending={contentPending}
          note="Khối nội dung ứng viên; nhãn xác nhận vùng nội dung cải tiến được sao chép nguyên văn."
          action={<Button onClick={() => onGo('contentReview')}>Kiểm tra nội dung cải tiến</Button>}
        />
      </div>

      <Card title="Trạng thái mô hình">
        <div className="flex flex-col gap-2">
          <div className="flex flex-wrap items-start gap-2 rounded-sm2 border border-lineSoft bg-[#fafcfe] px-2.5 py-2">
            <Images className="mt-[2px] h-4 w-4 shrink-0 text-brand-600" />
            <span className="w-[130px] shrink-0 text-base2 font-medium text-header">Mô hình ảnh</span>
            <ToneChip tone={imageModel.tone} icon={imageModel.tone === 'ok'
              ? <CheckCircle2 className="h-3.5 w-3.5" /> : <AlertTriangle className="h-3.5 w-3.5" />}>
              {imageModel.label}
            </ToneChip>
            <p className="min-w-[240px] flex-1 text-base2 text-body">{learning.modelStatus.image}</p>
          </div>
          <div className="flex flex-wrap items-start gap-2 rounded-sm2 border border-lineSoft bg-[#fafcfe] px-2.5 py-2">
            <Text className="mt-[2px] h-4 w-4 shrink-0 text-brand-600" />
            <span className="w-[130px] shrink-0 text-base2 font-medium text-header">Mô hình nội dung</span>
            <ToneChip tone={contentModel.tone} icon={contentModel.tone === 'ok'
              ? <CheckCircle2 className="h-3.5 w-3.5" /> : <AlertTriangle className="h-3.5 w-3.5" />}>
              {contentModel.label}
            </ToneChip>
            <p className="min-w-[240px] flex-1 text-base2 text-body">{learning.modelStatus.content}</p>
          </div>
          <p className="flex gap-1.5 text-xs2 leading-[16px] text-muted">
            <Info className="mt-[1px] h-3.5 w-3.5 shrink-0 text-brand-500" />
            Dữ liệu học ảnh và dữ liệu học nội dung được tách riêng và huấn luyện độc lập. Khi chưa đủ dữ liệu học,
            chương trình vẫn dùng quy tắc hiện tại và AI chỉ xác định vị trí nội dung — văn bản và ảnh gốc do chương
            trình sao chép.
          </p>
        </div>
      </Card>
    </>
  )
}

function LearningGroup({
  title,
  icon,
  total,
  labeled,
  pending,
  note,
  action,
}: {
  title: string
  icon: React.ReactNode
  total: number
  labeled: number
  pending: number
  note: string
  action: React.ReactNode
}) {
  const percent = total ? Math.round((labeled / total) * 100) : 0
  return (
    <Card title={title} actions={<span className="text-brand-500">{icon}</span>}>
      <div className="flex items-end gap-4">
        <div>
          <p className="text-[22px] font-semibold leading-[26px] text-header">{pending}</p>
          <p className="text-xs2 text-muted">đối tượng cần kiểm tra</p>
        </div>
        <div className="flex-1">
          <div className="flex items-center justify-between text-xs2">
            <span className="text-body">
              Đã gán nhãn <span className="font-semibold">{labeled}</span>/{total} đối tượng
            </span>
            <span className="font-semibold text-brand-600">{percent}%</span>
          </div>
          <ProgressBar className="mt-1" value={percent} />
        </div>
      </div>
      <p className="mt-2 text-xs2 leading-[16px] text-muted">{note}</p>
      <div className="mt-2">{action}</div>
    </Card>
  )
}

/* ------------------------------------------------------------------ Kiểm tra ảnh */

/** PROMPT-024R §20: target rectangle in slide-% space, straight from the authoritative Python DTO. */
function targetRect(c: ImageCandidate): { x: number; y: number; w: number; h: number } {
  const pct = c.targetBboxPct
  if (pct) return { x: pct.x, y: pct.y, w: pct.w, h: pct.h }
  return { x: c.bounds.x, y: c.bounds.y, w: c.bounds.w, h: c.bounds.h }
}

/**
 * PROMPT-027 §13/§15/§16: the FINAL Excel evidence region is a DIFFERENT geometry from the picture
 * candidate being labelled.  One ImprovementItem can own several After pictures plus caption/arrows, so
 * the region exported to Excel is generally LARGER than one picture.  Its geometry comes from Python's
 * authoritative ``ImprovementVisualRegion.bbox`` (the exact region production crops) already expressed as
 * slide percentages — React never derives it from ``candidate.bounds`` and never guesses from pixels.
 * Returns null when this picture does not belong to any region that would reach Excel.
 */
function evidenceRect(c: ImageCandidate): { x: number; y: number; w: number; h: number } | null {
  const pct = c.evidenceRegionBboxPct
  if (!pct || !(pct.w > 0) || !(pct.h > 0)) return null
  return { x: pct.x, y: pct.y, w: pct.w, h: pct.h }
}

/** Which overlay geometry the reviewer wants to see; "both" is the default so the Excel region is obvious. */
type OverlayMode = 'both' | 'evidence' | 'picture'

const OVERLAY_MODES: { key: OverlayMode; label: string }[] = [
  { key: 'evidence', label: 'Vùng xuất Excel' },
  { key: 'picture', label: 'Ảnh đang đánh giá' },
  { key: 'both', label: 'Cả hai' },
]

/** Human-readable renderer backend name for the preview (§4/§42) — never a filesystem path. */
function backendLabel(backend: string): string {
  if (backend === 'powerpoint') return 'PowerPoint'
  if (backend === 'libreoffice') return 'LibreOffice'
  if (backend === 'builtin') return 'Trình vẽ tích hợp'
  return backend || 'không rõ'
}

const MIN_LIST_W = 200
const MAX_LIST_W = 420

function ImageReview() {
  const s = useStore()
  const list = s.learning.images
  const index = list.length ? Math.min(Math.max(s.imageReviewIndex, 0), list.length - 1) : -1
  const cand = list[index]

  /* §22: horizontally resizable left panel (drag the divider). */
  const [listWidth, setListWidth] = useState(250)
  const dragRef = useRef<{ startX: number; startW: number } | null>(null)
  const startDrag = (e: React.PointerEvent<HTMLDivElement>) => {
    dragRef.current = { startX: e.clientX, startW: listWidth }
    e.currentTarget.setPointerCapture(e.pointerId)
  }
  const moveDrag = (e: React.PointerEvent<HTMLDivElement>) => {
    if (!dragRef.current) return
    const next = dragRef.current.startW + (e.clientX - dragRef.current.startX)
    setListWidth(Math.max(MIN_LIST_W, Math.min(MAX_LIST_W, next)))
  }
  const endDrag = () => { dragRef.current = null }

  /* §31: fit-to-viewport preview with zoom; geometry recomputed on container resize (§28/§37). */
  const viewportRef = useRef<HTMLDivElement>(null)
  const [viewport, setViewport] = useState({ w: 0, h: 0 })
  const [zoom, setZoom] = useState(1)
  /* PROMPT-027 §17: compact overlay toggle.  Declared with the other hooks and BEFORE the
     candidate-dependent early return below, so hook order stays stable (React #310 / §44). */
  const [overlayMode, setOverlayMode] = useState<OverlayMode>('both')
  useEffect(() => {
    const el = viewportRef.current
    if (!el) return
    const update = () => setViewport({ w: el.clientWidth, h: el.clientHeight })
    update()
    const observer = new ResizeObserver(update)
    observer.observe(el)
    return () => observer.disconnect()
  }, [])

  /* Compact report groups: the file name is shown once per report, never repeated on every item (§21). */
  const groups = useMemo(() => {
    const out: { key: string; managementNumber: string; fileName: string; items: { cand: ImageCandidate; idx: number }[] }[] = []
    list.forEach((c, idx) => {
      const key = `${c.managementNumber}|${c.sourceFile}`
      const last = out[out.length - 1]
      if (last && last.key === key) last.items.push({ cand: c, idx })
      else out.push({ key, managementNumber: c.managementNumber, fileName: c.sourceFile, items: [{ cand: c, idx }] })
    })
    return out
  }, [list])

  const selectedRef = useRef<HTMLButtonElement>(null)
  useEffect(() => {
    selectedRef.current?.scrollIntoView({ block: 'nearest' })
  }, [index])

  /* §36: below the xl breakpoint the layout stacks, so the list panel must not keep its desktop width. */
  const [isDesktop, setIsDesktop] = useState(() => window.matchMedia('(min-width: 1280px)').matches)
  useEffect(() => {
    const mq = window.matchMedia('(min-width: 1280px)')
    const onChange = (e: MediaQueryListEvent) => setIsDesktop(e.matches)
    mq.addEventListener('change', onChange)
    return () => mq.removeEventListener('change', onChange)
  }, [])

  if (!cand) return (
    <Card title="Kiểm tra ảnh cải tiến">
      <div className="flex flex-col items-start gap-2 text-base2 text-muted">
        <p>{s.connected ? 'Python backend chưa trả về ảnh ứng viên cần kiểm tra.' : 'Đang chờ kết nối Python backend.'}</p>
        <Button icon={<RefreshCw className="h-3.5 w-3.5" />} onClick={() => void s.refreshLearning()}>Làm mới trạng thái học</Button>
      </div>
    </Card>
  )

  const chosen = cand.userLabel !== 'UNLABELED'
  const rect = targetRect(cand)
  /* PROMPT-027 §13: TWO different review geometries.  ``rect`` is the picture currently being labelled;
     ``evidence`` is the item-scoped region production would export to Excel.  They are never conflated. */
  const evidence = evidenceRect(cand)
  /* §16/§18: the sharp (un-dimmed) window follows the FINAL Excel region when one exists — those are the
     pixels the reviewer must judge.  Blur/dim stays pure React visualization: the rendered slide bitmap,
     the ImprovementVisualRegion, the production crop and the Excel bytes are never modified here. */
  const focus = evidence ?? rect
  const clipTop = Math.max(0, focus.y)
  const clipRight = Math.max(0, 100 - (focus.x + focus.w))
  const clipBottom = Math.max(0, 100 - (focus.y + focus.h))
  const clipLeft = Math.max(0, focus.x)
  const showEvidence = overlayMode !== 'picture' && evidence !== null
  const showPicture = overlayMode !== 'evidence'
  const evidenceItemIndex = typeof cand.evidenceRegionItemIndex === 'number' && cand.evidenceRegionItemIndex >= 0
    ? cand.evidenceRegionItemIndex
    : cand.itemIndex
  const evidenceLabel = typeof evidenceItemIndex === 'number' && evidenceItemIndex >= 0
    ? `Mục #${evidenceItemIndex + 1} · Vùng xuất Excel`
    : 'Vùng xuất Excel'
  /* §4/§6/§41: honest backend reporting.  Only PowerPoint is pixel-faithful; any fallback backend gets a
     small NON-BLOCKING fidelity notice instead of silently pretending to match PowerPoint. */
  const previewBackend = cand.slidePreviewBackend || ''
  const previewFaithful = Boolean(cand.slidePreviewFaithful)
  const showFidelityNotice = Boolean(previewBackend) && !previewFaithful

  /* §25/§28: true authored position — the slide's own aspect ratio drives the display box, so the
     percentage highlight stays aligned at any panel size, DPI scale or zoom level. */
  const slideAspect = cand.slideWidth && cand.slideHeight
    ? cand.slideWidth / cand.slideHeight
    : cand.slidePreviewWidth && cand.slidePreviewHeight
      ? cand.slidePreviewWidth / cand.slidePreviewHeight
      : 16 / 9
  const pad = 24
  const availW = Math.max(80, viewport.w - pad * 2)
  const availH = Math.max(80, viewport.h - pad * 2)
  let dispW = availW
  let dispH = availW / slideAspect
  if (dispH > availH) { dispH = availH; dispW = availH * slideAspect }
  dispW *= zoom
  dispH *= zoom
  const itemLabel = typeof cand.itemIndex === 'number' && cand.itemIndex >= 0 ? `Mục #${cand.itemIndex + 1}` : `ảnh #${cand.pictureId}`

  const leftPanel = (
    <section
      className="card flex min-h-0 flex-col xl:shrink-0"
      style={isDesktop ? { width: listWidth } : undefined}
      aria-label="Đối tượng ảnh"
    >
      <header className="card-head">
        <h2 className="card-title mr-auto">Đối tượng ảnh</h2>
        <span className="text-xs2 text-muted">
          {list.filter((c) => c.userLabel === 'UNLABELED').length} chưa xác nhận · {list.length} tổng
        </span>
      </header>
      <div className="min-h-0 flex-1 overflow-y-auto xl:max-h-none max-h-[300px]">
        {groups.map((group) => (
          <div key={group.key}>
            <div
              className="sticky top-0 z-10 border-b border-lineSoft bg-[#f4f7fb] px-2 py-1 text-xxs text-muted"
              title={`${group.managementNumber} — ${group.fileName}`}
            >
              <span className="font-medium text-header">{group.managementNumber}</span>
              <span className="mx-1">·</span>
              <span className="align-middle">{group.fileName}</span>
            </div>
            <ul className="divide-y divide-lineSoft">
              {group.items.map(({ cand: c, idx }) => {
                const selected = idx === index
                return (
                  <li key={c.id}>
                    <button
                      type="button"
                      ref={selected ? selectedRef : undefined}
                      onClick={() => s.setImageReviewIndex(idx)}
                      title={`${c.sourceFile} — Slide ${c.slide} — ${itemLabel}`}
                      className={cn(
                        'flex w-full items-center gap-2 px-2 py-1.5 text-left',
                        selected ? 'bg-brand-50' : 'bg-white hover:bg-[#f6f9fd]',
                      )}
                    >
                      {c.src
                        ? <img src={c.src} alt="" className="h-[34px] w-[52px] shrink-0 rounded-[2px] border border-line object-cover" />
                        : <span className="flex h-[34px] w-[52px] shrink-0 items-center justify-center rounded-[2px] border border-line bg-[#f4f7fb] text-xxs text-muted">—</span>}
                      <span className="min-w-0 flex-1">
                        <span className="flex flex-wrap items-center gap-1 text-xs2">
                          <span className="font-medium text-header">Slide {c.slide}</span>
                          {typeof c.itemIndex === 'number' && c.itemIndex >= 0 ? (
                            <span className="rounded-[2px] bg-brand-50 px-1 font-medium text-brand-600">Mục #{c.itemIndex + 1}</span>
                          ) : (
                            <span className="text-muted">ảnh #{c.pictureId}</span>
                          )}
                        </span>
                        <span className="mt-0.5 block">
                          <ToneChip tone={c.userLabel === 'UNLABELED' ? 'muted' : c.labelPending ? 'warn' : 'ok'}>
                            {c.userLabel === 'UNLABELED' ? 'Chưa xác nhận' : `${labelName(c.userLabel)}${c.labelPending ? ' · chưa lưu' : ''}`}
                          </ToneChip>
                        </span>
                      </span>
                    </button>
                  </li>
                )
              })}
            </ul>
          </div>
        ))}
      </div>
    </section>
  )

  const centerPanel = (
    <section className="card flex min-h-0 min-w-0 flex-1 flex-col" aria-label="Slide ngữ cảnh">
      <header className="card-head gap-2">
        <h2 className="card-title">Slide {cand.slide}</h2>
        {typeof cand.itemIndex === 'number' && cand.itemIndex >= 0 && (
          <span className="rounded-[2px] bg-brand-50 px-1.5 py-[1px] text-xs2 font-medium text-brand-600">
            Mục #{cand.itemIndex + 1}
          </span>
        )}
        <span className="hidden min-w-0 flex-1 truncate text-xs2 text-muted sm:inline" title={cand.itemHeading || ''}>
          {cand.itemHeading || ''}
        </span>
        <div className="ml-auto flex items-center gap-1">
          {/* PROMPT-027 §17: compact overlay toggle — deliberately NOT a large toolbar. */}
          <div className="mr-1 flex items-center gap-[2px] rounded-sm2 border border-line bg-white/70 p-[2px]"
            role="group" aria-label="Chọn khung hiển thị trên slide">
            {OVERLAY_MODES.map(({ key, label }) => (
              <button
                key={key}
                type="button"
                onClick={() => setOverlayMode(key)}
                aria-pressed={overlayMode === key}
                className={cn(
                  'whitespace-nowrap rounded-[2px] px-1.5 py-[2px] text-xxs',
                  overlayMode === key ? 'bg-brand-500 font-semibold text-white' : 'text-muted hover:bg-[#f4f7fb]',
                )}
              >
                {label}
              </button>
            ))}
          </div>
          {/* §4/§42: which renderer actually produced this preview (safe name only, never a path). */}
          {previewBackend && (
            <span
              className="hidden shrink-0 rounded-[2px] border border-line bg-white/70 px-1.5 py-[1px] text-xxs text-muted lg:inline"
              title={`Backend đã render ảnh xem trước: ${previewBackend}`}
            >
              Render: {backendLabel(previewBackend)}
            </span>
          )}
          <Button variant="ghost" className="h-[24px] px-1.5" disabled={zoom <= 0.5}
            onClick={() => setZoom((z) => Math.max(0.5, z / 1.25))} aria-label="Thu nhỏ">
            <Minus className="h-3.5 w-3.5" />
          </Button>
          <span className="w-[42px] text-center font-mono text-xs2 text-body">{Math.round(zoom * 100)}%</span>
          <Button variant="ghost" className="h-[24px] px-1.5" disabled={zoom >= 4}
            onClick={() => setZoom((z) => Math.min(4, z * 1.25))} aria-label="Phóng to">
            <Plus className="h-3.5 w-3.5" />
          </Button>
          <Button variant="ghost" className="h-[24px] px-2" icon={<Maximize2 className="h-3.5 w-3.5" />}
            onClick={() => setZoom(1)}>
            Vừa khung
          </Button>
        </div>
      </header>
      {/* PROMPT-027 §6/§41: honest fallback notice — small and NON-BLOCKING.  Reduced preview fidelity must
          never fail the report, and a genuine PowerPoint render must not show this warning at all. */}
      {showFidelityNotice && (
        <p
          className="flex items-center gap-1.5 border-b border-warn-500/30 bg-warn-50 px-3 py-1 text-xxs text-warn-600"
          role="status"
        >
          <AlertTriangle className="h-3 w-3 shrink-0" />
          <span>
            Bản xem trước đơn giản — bố cục có thể khác PowerPoint (backend: {backendLabel(previewBackend)}).
            Kết quả Excel vẫn dùng đúng toạ độ PPTX.
          </span>
        </p>
      )}
      <div ref={viewportRef} className={cn('relative min-h-0 flex-1 overflow-auto bg-[#dbe3ec]', !cand.slidePreview && 'flex items-center justify-center')}>
        {cand.slidePreview ? (
          <div className="flex min-h-full min-w-full items-center justify-center p-3">
            {/* §23/§24: the FULL authored slide stays visible; non-target content is dimmed/blurred,
                the target keeps its TRUE authored position and stays sharp (§25). Review display only —
                evidence bytes are never modified (§30). */}
            <div className="relative shrink-0 overflow-hidden rounded-[2px] border border-line bg-white shadow-card"
              style={{ width: dispW, height: dispH }}>
              <img
                src={cand.slidePreview} alt="" draggable={false}
                className="absolute inset-0 h-full w-full select-none"
                style={{ filter: 'blur(2px) brightness(0.82) saturate(0.85)' }}
              />
              <img
                src={cand.slidePreview} alt="" draggable={false}
                className="absolute inset-0 h-full w-full select-none"
                style={{ clipPath: `inset(${clipTop}% ${clipRight}% ${clipBottom}% ${clipLeft}%)` }}
              />
              {/* PROMPT-027 §16: SECONDARY — the picture currently being classified/labelled.  Thin dashed
                  box, drawn first so the Excel region stays visually dominant.  Purely a highlight. */}
              {showPicture && (
                <div
                  data-testid="overlay-picture-candidate"
                  className="absolute rounded-[2px] border border-dashed border-amber-500"
                  style={{ left: `${rect.x}%`, top: `${rect.y}%`, width: `${rect.w}%`, height: `${rect.h}%` }}
                >
                  <span className="absolute bottom-0 left-0 whitespace-nowrap rounded-t-[2px] bg-amber-500 px-1 py-[1px] text-xxs font-medium text-white">
                    Ảnh #{cand.pictureId}
                  </span>
                </div>
              )}
              {/* PROMPT-027 §16: PRIMARY — the FINAL Excel evidence region of this candidate's improvement
                  item.  Solid, heavier border and the more prominent label, because this is the exact
                  geometry production crops into Excel (possibly several After pictures + annotation). */}
              {showEvidence && evidence && (
                <div
                  data-testid="overlay-evidence-region"
                  className="absolute rounded-[2px] border-2 border-brand-600 shadow-[0_0_0_1px_rgba(255,255,255,0.75)]"
                  style={{ left: `${evidence.x}%`, top: `${evidence.y}%`, width: `${evidence.w}%`, height: `${evidence.h}%` }}
                >
                  <span
                    className={cn(
                      'absolute left-0 whitespace-nowrap rounded-[2px] bg-brand-600 px-1 py-[1px] text-xxs font-semibold text-white',
                      evidence.y > 3 ? '-top-[18px]' : 'top-0',
                    )}
                  >
                    {evidenceLabel}
                  </span>
                </div>
              )}
            </div>
          </div>
        ) : cand.src ? (
          <div className="flex flex-col items-center gap-2 p-4">
            <img src={cand.src} alt={cand.nearbyText ?? ''} className="max-h-[420px] max-w-full rounded-[2px] border border-line object-contain" />
            <p className="max-w-[460px] text-center text-xs2 text-muted">
              Chưa render được toàn slide để làm ngữ cảnh — hiển thị ảnh trích xuất. Vị trí khung vẫn theo toạ độ gốc của slide.
            </p>
          </div>
        ) : (
          <p className="text-base2 text-muted">Preview không khả dụng.</p>
        )}
      </div>
      <div className="border-t border-lineSoft px-3 py-1.5 text-xxs leading-[15px] text-muted">
        <p className="flex flex-wrap items-center gap-x-3 gap-y-1">
          <span className="inline-flex items-center gap-1">
            <span className="inline-block h-[9px] w-[14px] rounded-[1px] border-2 border-brand-600" aria-hidden />
            <strong className="font-semibold text-body">Vùng xuất Excel</strong>
            {evidence ? `— ${evidenceLabel}` : '— ảnh này không thuộc vùng Sau cải tiến nào sẽ xuất Excel'}
          </span>
          <span className="inline-flex items-center gap-1">
            <span className="inline-block h-[9px] w-[14px] rounded-[1px] border border-dashed border-amber-500" aria-hidden />
            <strong className="font-semibold text-body">Ảnh đang đánh giá</strong> — Ảnh #{cand.pictureId}
          </span>
        </p>
        <p className="mt-0.5">
          Hai khung là HAI khái niệm khác nhau: khung nét đứt là một ảnh ứng viên, khung liền nét là vùng bằng
          chứng Sau cải tiến của cả Mục (có thể gồm nhiều ảnh + chú thích) mà Excel sẽ nhận. Vùng làm nét là vị trí
          thật trên slide gốc; phần còn lại chỉ được làm mờ để dễ quan sát và không ảnh hưởng dữ liệu xuất.
        </p>
      </div>
    </section>
  )

  const rightPanel = (
    <section className="card flex w-full shrink-0 flex-col xl:w-[330px]" aria-label="Thông tin và nhãn xác nhận">
      <header className="card-head">
        <h2 className="card-title mr-auto">Thông tin & nhãn xác nhận</h2>
        <span className="text-xs2 text-muted">Mục {index + 1}/{list.length}</span>
      </header>
      <div className="min-h-0 flex-1 overflow-y-auto p-3">
        <KeyValue label="Report nguồn"><span title={cand.sourceFile}>{cand.sourceFile}</span></KeyValue>
        <KeyValue label="Management Number">{cand.managementNumber}</KeyValue>
        <KeyValue label="Slide / Mục">
          Slide {cand.slide}
          {typeof cand.itemIndex === 'number' && cand.itemIndex >= 0 && ` · Mục #${cand.itemIndex + 1}`}
        </KeyValue>
        {cand.itemHeading && (
          <KeyValue label="Tiêu đề mục"><span title={cand.itemHeading}>{cand.itemHeading}</span></KeyValue>
        )}
        <KeyValue label="Dự đoán">{cand.decision}</KeyValue>
        <KeyValue label="Độ tin cậy">
          {CONFIDENCE_LABEL[cand.confidenceBand]} <span className="text-muted">({cand.confidence.toFixed(2)})</span>
        </KeyValue>
        <KeyValue label="Vào Excel">
          {cand.excelEligible ? 'Đủ điều kiện'
            : <span className="text-warn-600">Chưa đủ điều kiện — {cand.eligibilityReason}</span>}
        </KeyValue>

        {/* PROMPT-027 §13/§42: the two review geometries and the renderer that produced the preview, shown
            as plain numbers so Windows acceptance can compare UI, logs and the exported crop. */}
        <div className="mt-2.5 border-t border-lineSoft pt-2">
          <p className="mb-1 text-base2 font-medium text-header">Vùng xuất Excel (bằng chứng Sau)</p>
          {evidence && cand.evidenceRegionBbox ? (
            <>
              <KeyValue label="Thuộc Mục">
                {evidenceLabel}
                {typeof cand.evidenceRegionPictureCount === 'number' && cand.evidenceRegionPictureCount > 0 &&
                  ` · ${cand.evidenceRegionPictureCount} ảnh`}
              </KeyValue>
              <KeyValue label="Toạ độ EMU">
                <span className="font-mono text-xxs">
                  {cand.evidenceRegionBbox.x}, {cand.evidenceRegionBbox.y} ·
                  {' '}{cand.evidenceRegionBbox.width}×{cand.evidenceRegionBbox.height}
                </span>
              </KeyValue>
              <KeyValue label="% slide">
                <span className="font-mono text-xxs">
                  {evidence.x.toFixed(2)}, {evidence.y.toFixed(2)} · {evidence.w.toFixed(2)}×{evidence.h.toFixed(2)}
                </span>
              </KeyValue>
            </>
          ) : (
            <p className="text-xs2 text-warn-600">
              Ảnh này hiện KHÔNG nằm trong vùng Sau cải tiến nào sẽ được xuất Excel (khung nét liền không hiển thị).
            </p>
          )}
          <KeyValue label="Ảnh đang đánh giá">
            <span className="font-mono text-xxs">
              Ảnh #{cand.pictureId} · {rect.x.toFixed(2)}, {rect.y.toFixed(2)} · {rect.w.toFixed(2)}×{rect.h.toFixed(2)}%
            </span>
          </KeyValue>
          <KeyValue label="Render xem trước">
            {previewBackend ? (
              <span>
                {backendLabel(previewBackend)}
                {previewFaithful
                  ? <span className="text-ok-600"> · trung thực PowerPoint</span>
                  : <span className="text-warn-600"> · bản đơn giản, bố cục có thể khác</span>}
              </span>
            ) : '—'}
          </KeyValue>
        </div>

        <p className="mb-1 mt-2.5 text-base2 font-medium text-header">Lý do</p>
        <ul className="max-h-[110px] list-disc overflow-y-auto pl-5 text-xs2 leading-[17px] text-body">
          {cand.evidence.map((e) => <li key={e}>{e}</li>)}
        </ul>
        {cand.nearbyText && <p className="mt-1.5 text-xs2 text-muted">Chữ gần ảnh: “{cand.nearbyText}”</p>}

        <div className="mt-3 border-t border-lineSoft pt-2.5">
          <p className="mb-1.5 text-base2 font-medium text-header">Chọn nhãn xác nhận</p>
          <div className="grid grid-cols-2 gap-1.5">
            {IMAGE_LABELS.map((l) => (
              <Button key={l.key} active={cand.userLabel === l.key} disabled={s.locked}
                className="justify-center"
                onClick={() => s.labelImage(cand.id, l.key)}>
                {l.label}
              </Button>
            ))}
          </div>

          <p className="mb-1 mt-3 text-base2 font-medium text-header">Ghi chú</p>
          <TextInput
            value={cand.note}
            disabled={s.locked}
            maxLength={1000}
            placeholder="Ghi chú cho đối tượng này (không bắt buộc)"
            onChange={(e) => s.setNote('image', cand.id, e.target.value)}
          />

          <div className="mt-3 flex flex-wrap items-center gap-2">
            <Button icon={<ChevronLeft className="h-3.5 w-3.5" />} disabled={index === 0}
              onClick={() => s.setImageReviewIndex(index - 1)}>
              Mục trước
            </Button>
            <Button disabled={index >= list.length - 1} onClick={() => s.setImageReviewIndex(index + 1)}>
              Mục tiếp
              <ChevronRight className="h-3.5 w-3.5" />
            </Button>
          </div>
          <Button
            variant="primary"
            className="mt-2 w-full justify-center"
            icon={<Save className="h-3.5 w-3.5" />}
            disabled={!chosen || s.locked}
            onClick={() => void s.saveImageCandidate(cand.id, cand.userLabel, cand.note)}
          >
            Lưu nhãn & ghi chú
          </Button>
          {chosen && (
            <span className={cn('mt-2 inline-flex items-center gap-1.5 text-xs2', cand.labelPending ? 'text-warn-600' : 'text-ok-600')}>
              {cand.labelPending ? <AlertTriangle className="h-3.5 w-3.5" /> : <CheckCircle2 className="h-3.5 w-3.5" />}
              {cand.labelPending ? `Nhãn “${labelName(cand.userLabel)}” chưa lưu.` : `Nhãn “${labelName(cand.userLabel)}” đã lưu cục bộ.`}
            </span>
          )}
        </div>
      </div>
    </section>
  )

  /* §20/§35: three horizontal regions filling one viewport on desktop; each panel scrolls independently.
     §36: below xl the preview + info stack under a compact list instead of squeezing three columns. */
  return (
    <div className="flex min-h-0 flex-col gap-3 xl:h-[calc(100dvh-132px)] xl:min-h-[540px]">
      <div className="flex min-h-0 flex-1 flex-col gap-3 xl:flex-row xl:gap-0">
        {leftPanel}
        <div
          className="hidden w-[6px] shrink-0 cursor-col-resize touch-none items-stretch xl:flex"
          role="separator"
          aria-orientation="vertical"
          aria-label="Kéo để thay đổi độ rộng danh sách ảnh"
          onPointerDown={startDrag}
          onPointerMove={moveDrag}
          onPointerUp={endDrag}
          onPointerCancel={endDrag}
        >
          <div className="mx-auto w-[2px] rounded bg-line transition-colors hover:bg-brand-400" />
        </div>
        {centerPanel}
        {rightPanel}
      </div>
    </div>
  )
}

/* ------------------------------------------------------------------ Kiểm tra nội dung */

function ContentReview() {
  const s = useStore()
  const list = s.learning.contents
  const index = list.length ? Math.min(Math.max(s.contentReviewIndex, 0), list.length - 1) : -1
  const cand = list[index]

  if (!cand) return (
    <Card title="Kiểm tra nội dung cải tiến">
      <div className="flex flex-col items-start gap-2 text-base2 text-muted">
        <p>{s.connected ? 'Python backend chưa trả về khối nội dung cần kiểm tra.' : 'Đang chờ kết nối Python backend.'}</p>
        <Button icon={<RefreshCw className="h-3.5 w-3.5" />} onClick={() => void s.refreshLearning()}>Làm mới trạng thái học</Button>
      </div>
    </Card>
  )

  return (
    <div className="grid grid-cols-1 gap-3 xl:grid-cols-[300px_1fr]">
      <Card title="Vùng nội dung" actions={<span className="text-xs2 text-muted">{list.filter((c) => c.userLabel === 'UNLABELED').length} chưa xác nhận · {list.length} tổng</span>} dense>
        <ul className="divide-y divide-lineSoft">
          {list.map((c, i) => (
            <li key={c.id}>
              <button
                type="button"
                onClick={() => s.setContentReviewIndex(i)}
                className={cn('flex w-full flex-col gap-0.5 px-2.5 py-2 text-left', i === index ? 'bg-brand-50' : 'bg-white hover:bg-[#f6f9fd]')}
              >
                <span className="truncate text-base2 font-medium text-header">{c.sourceFile}</span>
                <span className="text-xs2 text-muted">
                  Slide {c.slide} · khối #{c.blockId} · {c.section}
                </span>
                <span className="truncate text-xs2 text-body">{c.text.split('\n')[0]}</span>
                <span className="mt-0.5">
                  <ToneChip tone={c.userLabel === 'UNLABELED' ? 'muted' : c.labelPending ? 'warn' : 'ok'}>
                    {c.userLabel === 'UNLABELED' ? 'Chưa xác nhận' : `${labelName(c.userLabel)}${c.labelPending ? ' · chưa lưu' : ''}`}
                  </ToneChip>
                </span>
              </button>
            </li>
          ))}
        </ul>
      </Card>

      <div className="flex min-w-0 flex-col gap-3">
        <Card
          title="Nguyên văn nội dung"
          actions={<span className="text-xs2 text-muted">{cand.nearestTitle}</span>}
        >
          <div className="max-h-[240px] overflow-auto rounded-sm2 border border-line bg-[#fbfcfe] px-2.5 py-2">
            <pre className="whitespace-pre-wrap break-words font-sans text-base2 leading-[20px] text-body">
              {cand.text}
            </pre>
          </div>
          <p className="mt-1.5 flex gap-1.5 text-xs2 text-muted">
            <Info className="mt-[1px] h-3.5 w-3.5 shrink-0 text-brand-500" />
            Không tóm tắt hoặc viết lại nội dung nguồn: phần nội dung cải tiến được sao chép đầy đủ và chỉ loại bỏ phần
            “Xử lý tạm thời”.
          </p>
        </Card>

        <Card
          title="Thông tin & nhãn xác nhận"
          actions={
            <span className="text-xs2 text-muted">
              Mục {index + 1}/{list.length}
            </span>
          }
        >
          <div className="grid grid-cols-1 gap-x-6 xl:grid-cols-2">
            <div>
              <KeyValue label="Report nguồn">{cand.sourceFile}</KeyValue>
              <KeyValue label="Management Number">{cand.managementNumber}</KeyValue>
              <KeyValue label="Slide">{cand.slide}</KeyValue>
              <KeyValue label="Khối">#{cand.blockId}</KeyValue>
              <KeyValue label="Mục gần nhất">{cand.section}</KeyValue>
              <KeyValue label="Dự đoán">{cand.decision}</KeyValue>
              <KeyValue label="Độ tin cậy">
                {CONFIDENCE_LABEL[cand.confidenceBand]}{' '}
                <span className="text-muted">({cand.confidence.toFixed(2)})</span>
              </KeyValue>
            </div>
            <div>
              <p className="mb-1 text-base2 font-medium text-header">Lý do</p>
              <ul className="list-disc pl-5 text-base2 text-body">
                {cand.evidence.map((e) => (
                  <li key={e}>{e}</li>
                ))}
              </ul>
            </div>
          </div>

          <div className="mt-3 border-t border-lineSoft pt-2.5">
            <p className="mb-1.5 text-base2 font-medium text-header">Chọn nhãn xác nhận</p>
            <div className="flex flex-wrap gap-2">
              {[
                { key: 'IMPROVEMENT_CONTENT', label: 'Nội dung cải tiến' },
                { key: 'EXCLUDE_CONTENT', label: 'Không lấy' },
              ].map((l) => (
                <Button key={l.key} active={cand.userLabel === l.key} disabled={s.locked}
                  onClick={() => s.labelContent(cand.id, l.key)}>
                  {l.label}
                </Button>
              ))}
            </div>

            <p className="mb-1 mt-3 text-base2 font-medium text-header">Ghi chú</p>
            <TextInput
              value={cand.note}
              disabled={s.locked}
              maxLength={1000}
              placeholder="Ghi chú cho vùng nội dung này (không bắt buộc)"
              onChange={(e) => s.setNote('content', cand.id, e.target.value)}
            />

            <div className="mt-3 flex flex-wrap items-center gap-2">
              <Button disabled={index === 0} onClick={() => s.setContentReviewIndex(index - 1)}>
                Mục trước
              </Button>
              <Button disabled={index >= list.length - 1} onClick={() => s.setContentReviewIndex(index + 1)}>
                Mục tiếp
              </Button>
              <Button variant="primary" icon={<Save className="h-3.5 w-3.5" />}
                disabled={cand.userLabel === 'UNLABELED' || s.locked}
                onClick={() => void s.saveContentCandidate(cand.id, cand.userLabel, cand.note)}>
                Lưu nhãn & ghi chú
              </Button>
              {cand.userLabel !== 'UNLABELED' && (
                <span className={cn('inline-flex items-center gap-1.5 text-xs2', cand.labelPending ? 'text-warn-600' : 'text-ok-600')}>
                  <CheckCircle2 className="h-3.5 w-3.5" />
                  {cand.labelPending ? `Nhãn “${labelName(cand.userLabel)}” chưa lưu.` : `Nhãn “${labelName(cand.userLabel)}” đã lưu cục bộ.`}
                </span>
              )}
            </div>
          </div>
        </Card>
      </div>
    </div>
  )
}

/* ------------------------------------------------------------------ Mô hình & dữ liệu */

function ModelAndData() {
  const s = useStore()
  const { learning, training } = s
  const result = learning.excelLastResult
  const imageModel = modelSummary(learning.modelStatus.image)
  const contentModel = modelSummary(learning.modelStatus.content)

  return (
    <>
      <Card title="Mô hình học">
        <div className="flex flex-col gap-2">
          <div className="flex flex-wrap items-center gap-2">
            <Images className="h-4 w-4 text-brand-600" />
            <span className="w-[130px] text-base2 font-medium text-header">Mô hình ảnh</span>
            <ToneChip tone={imageModel.tone} icon={imageModel.tone === 'ok'
              ? <CheckCircle2 className="h-3.5 w-3.5" /> : <AlertTriangle className="h-3.5 w-3.5" />}>
              {imageModel.label}
            </ToneChip>
            <span className="min-w-[240px] flex-1 text-base2 text-body">{learning.modelStatus.image}</span>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <Text className="h-4 w-4 text-brand-600" />
            <span className="w-[130px] text-base2 font-medium text-header">Mô hình nội dung</span>
            <ToneChip tone={contentModel.tone} icon={contentModel.tone === 'ok'
              ? <CheckCircle2 className="h-3.5 w-3.5" /> : <AlertTriangle className="h-3.5 w-3.5" />}>
              {contentModel.label}
            </ToneChip>
            <span className="min-w-[240px] flex-1 text-base2 text-body">{learning.modelStatus.content}</span>
          </div>
          <div className="flex flex-wrap items-center gap-2 border-t border-lineSoft pt-2.5">
            <Button
              variant="primary"
              icon={training ? <Loader2 className="h-3.5 w-3.5 animate-spin" /> : <Cpu className="h-3.5 w-3.5" />}
              disabled={s.locked}
              onClick={s.trainModels}
            >
              {training ? 'Đang cập nhật mô hình học…' : 'Cập nhật mô hình học'}
            </Button>
            <span className="text-xs2 text-muted">
              {learning.trainingMessage || 'Hai mô hình được huấn luyện độc lập; trạng thái và kết quả được lấy từ Python backend.'}
            </span>
          </div>
        </div>
      </Card>

      <Card title="Dữ liệu học">
        <div className="flex flex-col gap-2">
          <div className="flex flex-wrap items-center gap-2">
            <Button icon={<FolderOpen className="h-3.5 w-3.5" />} onClick={s.openLearningFolder}>
              Mở thư mục dữ liệu học
            </Button>
            <Button icon={<Upload className="h-3.5 w-3.5" />} onClick={s.exportLearningData}>
              Xuất dữ liệu học
            </Button>
            <span className="text-xs2 text-muted">Thư mục cục bộ do Python backend quản lý</span>
          </div>
          <p className="text-xs2 leading-[16px] text-muted">
            Nhãn được lưu cục bộ trong thư mục learning_data cạnh chương trình và được giữ lại khi cập nhật phiên bản.
            Dữ liệu học ảnh (image_labels.jsonl) và dữ liệu học nội dung (content_labels.jsonl) nằm riêng biệt.
          </p>
        </div>
      </Card>

      <Card title="Cập nhật Excel từ nhãn đã lưu">
        <div className="flex flex-col gap-2">
          <div className="flex flex-wrap items-center gap-2">
            <Button
              variant="primary"
              icon={<FileSpreadsheet className="h-3.5 w-3.5" />}
              disabled={s.locked || learning.excelPending === 0}
              onClick={() => s.applyLabelsToExcel(true)}
            >
              Thử lại cập nhật Excel đang chờ
            </Button>
            <span className="text-xs2 text-muted">
              {learning.excelPending > 0
                ? `${learning.excelPending} mục nhãn đã lưu đang chờ thao tác Excel được thử lại.`
                : 'Nhãn mới được áp dụng ngay khi lưu; hiện không có thao tác Excel nào đang chờ.'}
            </span>
          </div>

          {result.kind === 'locked' && (
            <div className="rounded-sm2 border border-warn-500/30 bg-warn-50 px-2.5 py-2">
              <p className="flex items-start gap-1.5 text-base2 text-warn-600">
                <ShieldAlert className="mt-[1px] h-4 w-4 shrink-0" />
                {result.message}
              </p>
              <div className="mt-2 flex flex-wrap items-center gap-2">
                <Button variant="primary" disabled={s.locked}
                  icon={<RefreshCw className="h-3.5 w-3.5" />} onClick={() => s.applyLabelsToExcel(true)}>
                  Thử lại
                </Button>
                <Button onClick={() => void s.dismissExcelResult()}>Để sau</Button>
                <span className="text-xs2 text-muted">
                  Nhãn không bị mất; chỉ bước cập nhật Excel được thực hiện lại.
                </span>
              </div>
            </div>
          )}

          {result.kind === 'ok' && result.message && (
            <div className="flex items-start gap-1.5 rounded-sm2 border border-ok-500/25 bg-ok-50 px-2.5 py-2 text-base2 text-ok-600">
              <CheckCircle2 className="mt-[1px] h-4 w-4 shrink-0" />
              {result.message}
            </div>
          )}

          <p className="text-xs2 leading-[16px] text-muted">
            Ghi vào file kết quả bằng cách ghi tạm rồi thay thế nguyên tử (os.replace) và không bao giờ ghi đè file
            kiểm chứng gốc.
          </p>
        </div>
      </Card>
    </>
  )
}

function modelSummary(status: string): { tone: 'ok' | 'warn' | 'error' | 'info'; label: string } {
  if (status.startsWith('Đã huấn luyện')) return { tone: 'ok', label: 'Đã huấn luyện' }
  if (status.toLowerCase().includes('không khả dụng')) return { tone: 'error', label: 'Không khả dụng' }
  if (status.startsWith('Đang')) return { tone: 'info', label: 'Đang cập nhật' }
  return { tone: 'warn', label: 'Chưa huấn luyện' }
}

function labelName(key: string) {
  return (
    {
      AFTER: 'Sau cải tiến',
      BEFORE: 'Trước cải tiến',
      CONTROL: 'Kiểm tra / Kiểm soát',
      IGNORE: 'Không lấy',
      IMPROVEMENT_CONTENT: 'Nội dung cải tiến',
      EXCLUDE_CONTENT: 'Không lấy',
    } as Record<string, string>
  )[key] ?? key
}
