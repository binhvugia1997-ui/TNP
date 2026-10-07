import { useState } from 'react'
import {
  AlertTriangle,
  BarChart3,
  CheckCircle2,
  Cpu,
  Database,
  FileSpreadsheet,
  FolderOpen,
  Images,
  Info,
  Loader2,
  RefreshCw,
  Save,
  ShieldAlert,
  Text,
  Upload,
} from 'lucide-react'
import { Button, Card, KeyValue, ProgressBar, TextInput, ToneChip } from '../components/ui'
import { useStore, type LearningState } from '../state/store'
import { CONFIDENCE_LABEL } from '../types'
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

function ImageReview() {
  const s = useStore()
  const list = s.learning.images
  const index = list.length ? Math.min(Math.max(s.imageReviewIndex, 0), list.length - 1) : -1
  const cand = list[index]

  if (!cand) return (
    <Card title="Kiểm tra ảnh cải tiến">
      <div className="flex flex-col items-start gap-2 text-base2 text-muted">
        <p>{s.connected ? 'Python backend chưa trả về ảnh ứng viên cần kiểm tra.' : 'Đang chờ kết nối Python backend.'}</p>
        <Button icon={<RefreshCw className="h-3.5 w-3.5" />} onClick={() => void s.refreshLearning()}>Làm mới trạng thái học</Button>
      </div>
    </Card>
  )

  const labelText = (key: string) => cand.userLabel === key
  const chosen = cand.userLabel !== 'UNLABELED'

  return (
    <div className="grid grid-cols-1 gap-3 xl:grid-cols-[300px_1fr]">
      <Card title="Đối tượng ảnh" actions={<span className="text-xs2 text-muted">{list.filter((c) => c.userLabel === 'UNLABELED').length} chưa xác nhận · {list.length} tổng</span>} dense>
        <ul className="divide-y divide-lineSoft">
          {list.map((c, i) => (
            <li key={c.id}>
              <button
                type="button"
                onClick={() => s.setImageReviewIndex(i)}
                className={cn('flex w-full items-start gap-2 px-2.5 py-2 text-left', i === index ? 'bg-brand-50' : 'bg-white hover:bg-[#f6f9fd]')}
              >
                <img src={c.src} alt="" className="h-[38px] w-[58px] shrink-0 rounded-[2px] border border-line object-cover" />
                <span className="min-w-0 flex-1">
                  <span className="block truncate text-base2 font-medium text-header">{c.sourceFile}</span>
                  <span className="block text-xs2 text-muted">
                    Slide {c.slide} · ảnh #{c.pictureId}
                  </span>
                  <span className="mt-0.5 block">
                    <ToneChip tone={c.userLabel === 'UNLABELED' ? 'muted' : c.labelPending ? 'warn' : 'ok'}>
                      {c.userLabel === 'UNLABELED' ? 'Chưa xác nhận' : `${labelName(c.userLabel)}${c.labelPending ? ' · chưa lưu' : ''}`}
                    </ToneChip>
                  </span>
                </span>
              </button>
            </li>
          ))}
        </ul>
      </Card>

      <div className="flex min-w-0 flex-col gap-3">
        <Card title="Ảnh xem trước" actions={<span className="text-xs2 text-muted">{cand.sourceFile}</span>}>
          <div className="relative aspect-[16/9] w-full overflow-hidden rounded-sm2 border border-line bg-[#f4f7fb]">
            <div className="absolute inset-x-0 top-0 flex h-[26px] items-center gap-2 border-b border-line bg-white px-2 text-xs2 text-muted">
              <span className="font-medium text-header">Slide {cand.slide}</span>
              {typeof cand.itemIndex === 'number' && cand.itemIndex >= 0 && (
                <>
                  <span>·</span>
                  <span className="rounded-[2px] bg-brand-50 px-1 font-medium text-brand-600">
                    Mục #{cand.itemIndex + 1}
                  </span>
                </>
              )}
              <span>·</span>
              <span>{cand.sourceFile}</span>
              <span className="ml-auto">Ảnh #{cand.pictureId}</span>
            </div>
            <div
              className="absolute overflow-hidden rounded-[2px] border-2 border-dashed border-brand-500 shadow-sm"
              style={{
                left: `${cand.bounds.x}%`,
                top: `${cand.bounds.y}%`,
                width: `${cand.bounds.w}%`,
                height: `${cand.bounds.h}%`,
              }}
            >
              {cand.src ? <img src={cand.src} alt={cand.nearbyText ?? ''} className="h-full w-full object-cover" />
                : <span className="flex h-full items-center justify-center text-xs2 text-muted">Preview không khả dụng</span>}
            </div>
            <span
              className="absolute rounded-[2px] bg-brand-500 px-1 py-[1px] text-xxs font-medium text-white"
              style={{ left: `${cand.bounds.x}%`, top: `calc(${cand.bounds.y}% - 16px)` }}
            >
              Ảnh #{cand.pictureId} · {cand.decision}
            </span>
          </div>
          <p className="mt-1.5 text-xs2 text-muted">
            Preview có kiểm soát do Python tạo; khung nét đứt biểu thị vị trí trên slide. Kết quả Excel vẫn tuân theo các
            quy tắc eligibility và bằng chứng ngữ nghĩa phía Python.
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
              <KeyValue label="Dự đoán">{cand.decision}</KeyValue>
              <KeyValue label="Độ tin cậy">
                <span className="inline-flex items-center gap-1.5">
                  {CONFIDENCE_LABEL[cand.confidenceBand]}
                  <span className="text-muted">({cand.confidence.toFixed(2)})</span>
                </span>
              </KeyValue>
              <KeyValue label="Vào Excel">
                {cand.excelEligible ? (
                  'Đủ điều kiện'
                ) : (
                  <span className="text-warn-600">Chưa đủ điều kiện — {cand.eligibilityReason}</span>
                )}
              </KeyValue>
            </div>
            <div>
              <p className="mb-1 text-base2 font-medium text-header">Lý do</p>
              <ul className="list-disc pl-5 text-base2 text-body">
                {cand.evidence.map((e) => (
                  <li key={e}>{e}</li>
                ))}
              </ul>
              {cand.nearbyText && <p className="mt-2 text-xs2 text-muted">Chữ gần ảnh: “{cand.nearbyText}”</p>}
            </div>
          </div>

          <div className="mt-3 border-t border-lineSoft pt-2.5">
            <p className="mb-1.5 text-base2 font-medium text-header">Chọn nhãn xác nhận</p>
            <div className="flex flex-wrap gap-2">
              {[
                { key: 'AFTER', label: 'Sau cải tiến' },
                { key: 'BEFORE', label: 'Trước cải tiến' },
                { key: 'CONTROL', label: 'Kiểm tra / Kiểm soát' },
                { key: 'IGNORE', label: 'Không lấy' },
              ].map((l) => (
                <Button key={l.key} active={labelText(l.key)} disabled={s.locked} onClick={() => s.labelImage(cand.id, l.key)}>
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
              <Button disabled={index === 0} onClick={() => s.setImageReviewIndex(index - 1)}>
                Mục trước
              </Button>
              <Button disabled={index >= list.length - 1} onClick={() => s.setImageReviewIndex(index + 1)}>
                Mục tiếp
              </Button>
              <Button
                variant="primary"
                icon={<Save className="h-3.5 w-3.5" />}
                disabled={!chosen || s.locked}
                onClick={() => void s.saveImageCandidate(cand.id, cand.userLabel, cand.note)}
              >
                Lưu nhãn & ghi chú
              </Button>
              {chosen && (
                <span className={cn('inline-flex items-center gap-1.5 text-xs2', cand.labelPending ? 'text-warn-600' : 'text-ok-600')}>
                  {cand.labelPending ? <AlertTriangle className="h-3.5 w-3.5" /> : <CheckCircle2 className="h-3.5 w-3.5" />}
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
