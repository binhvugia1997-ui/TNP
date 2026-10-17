import { AlertTriangle, CheckCircle2, Info, Loader2, RefreshCw, Search, Trash2 } from 'lucide-react'
import { Button, Card, Radio, TextInput, ToneChip } from '../ui'
import { useStore } from '../../state/store'
import { EXCEL_STATE_LABEL, type ExcelKind, type ExcelKindState } from '../../types'

/**
 * PROMPT-040: one remembered Excel file. The value is owned by Python (`config/config.json`); this
 * panel only shows it, offers “Đổi file…” and “Xóa lựa chọn”, and explains a missing/invalid default
 * instead of silently switching to another file or re-opening the picker.
 */
function ExcelDefaultField({
  kind,
  placeholder,
  value,
  state,
  disabled,
  onManual,
  onPick,
  onClear,
}: {
  kind: ExcelKind
  placeholder: string
  value: string
  state: ExcelKindState
  disabled: boolean
  onManual: (value: string) => void
  onPick: () => void
  onClear: () => void
}) {
  const fallbackLabel = kind === 'template' ? 'File kiểm chứng Excel' : 'File kết quả Excel'
  const label = state.label || fallbackLabel
  const ok = state.state === 'ok'
  const pending = state.state === 'unconfigured' || state.state === 'checking'
  const tone = ok ? 'ok' as const : pending ? 'muted' as const : 'warn' as const
  const icon = ok ? <CheckCircle2 className="h-3.5 w-3.5" />
    : state.state === 'checking' ? <Loader2 className="h-3.5 w-3.5 animate-spin" />
      : pending ? undefined : <AlertTriangle className="h-3.5 w-3.5" />

  return (
    <div className="flex flex-col gap-0.5">
      <div className="flex items-center gap-2">
        <label className="w-[132px] shrink-0 text-base2 text-body">{label}</label>
        <TextInput
          value={value}
          disabled={disabled}
          placeholder={placeholder}
          onChange={(event) => onManual(event.target.value)}
        />
        <Button disabled={disabled} aria-label={`Chọn ${label}`} onClick={onPick}>
          {state.path ? 'Đổi file…' : 'Chọn…'}
        </Button>
        <Button
          variant="ghost"
          disabled={disabled || !state.path}
          aria-label={`Xóa lựa chọn ${label}`}
          onClick={onClear}
        >
          <Trash2 className="h-3.5 w-3.5" />
          Xóa lựa chọn
        </Button>
      </div>

      <div className="flex flex-wrap items-center gap-x-2 gap-y-0.5 pl-[136px] text-xs2">
        <span className="text-muted">Excel mặc định:</span>
        {state.path ? (
          <>
            <span className="max-w-[320px] truncate font-mono text-body" title={state.path}>
              {state.name || state.path}
            </span>
            <ToneChip tone={tone} icon={icon}>{EXCEL_STATE_LABEL[state.state]}</ToneChip>
            <span className="min-w-0 flex-1 truncate text-muted" title={state.path}>{state.folder}</span>
          </>
        ) : (
          <span className="text-muted">chưa chọn — chọn một lần, các lần chạy sau dùng lại tự động</span>
        )}
      </div>

      {!ok && !pending && state.message && (
        <p role="status" className="pl-[136px] text-xs2 text-warn-600">{state.message}</p>
      )}
      {state.path !== '' && state.persisted === false && (
        <p role="status" className="pl-[136px] text-xs2 text-warn-600">
          Lựa chọn dùng được trong phiên này nhưng chưa ghi được vào config.json — kiểm tra quyền ghi thư mục config.
        </p>
      )}
    </div>
  )
}

export function SourcePanel() {
  const s = useStore()
  const disabled = s.locked

  return (
    <Card
      title="Nguồn dữ liệu"
      actions={<span className="text-xs2 text-muted">{disabled ? 'Đang xử lý — không thể thay đổi nguồn dữ liệu' : 'Chọn đường dẫn bằng hộp thoại Windows hoặc nhập trực tiếp.'}</span>}
    >
      <div className="flex flex-col gap-1.5">
        <div className="flex items-center gap-2">
          <label className="w-[132px] shrink-0 text-base2 text-body">Thư mục báo cáo PPTX</label>
          <TextInput
            value={s.folder}
            disabled={disabled}
            placeholder="Chọn thư mục báo cáo"
            onChange={(e) => s.setFolder(e.target.value)}
          />
          <Button disabled={disabled} onClick={s.chooseReportFolder}>Chọn…</Button>
          <Button disabled={disabled} onClick={s.choosePptxFile}>Chọn một PPTX…</Button>
          <Button variant="primary" disabled={disabled || s.scanBusy} icon={<Search className="h-3.5 w-3.5" />} onClick={s.scan}>
            {s.scanBusy ? 'Đang quét…' : 'Quét'}
          </Button>
        </div>
        <div className="pl-[136px] text-xs2 text-muted">
          {s.inputFileName ? `File được chọn: ${s.inputFileName} · ` : ''}
          {s.scanned && !s.scanDirty ? s.scanMessage : s.scanDirty ? 'Nguồn hoặc tùy chọn đã đổi — hãy quét lại trước khi xử lý.' : s.scanMessage}
        </div>

        <div className="grid grid-cols-1 gap-y-1.5">
          <ExcelDefaultField
            kind="template"
            placeholder="Chọn file kiểm chứng .xlsx"
            value={s.template}
            state={s.excel.template}
            disabled={disabled}
            onManual={s.setTemplate}
            onPick={s.chooseTemplateFile}
            onClear={() => s.clearExcel('template')}
          />
          <ExcelDefaultField
            kind="output"
            placeholder="Chọn hoặc nhập file kết quả .xlsx"
            value={s.output}
            state={s.excel.output}
            disabled={disabled}
            onManual={s.setOutput}
            onPick={s.chooseOutputFile}
            onClear={() => s.clearExcel('output')}
          />
        </div>

        <div className="mt-0.5 flex flex-wrap items-center gap-x-4 gap-y-1 border-t border-lineSoft pt-1.5">
          <span className="inline-flex items-center gap-1 text-xs2 text-muted">
            <Info className="h-3.5 w-3.5 text-brand-500" />
            Không ghi đè file kiểm chứng gốc
          </span>

          <span className="ml-auto flex flex-wrap items-center gap-x-3 gap-y-1">
            <span className="text-base2 text-body">Bộ lọc thời gian</span>
            <Radio name="period" checked={s.periodMode === 'auto'} onChange={() => s.setPeriodMode('auto')}
              label="Tự động theo file Excel" disabled={disabled} />
            <Radio name="period" checked={s.periodMode === 'all'} onChange={() => s.setPeriodMode('all')}
              label="Tất cả" disabled={disabled} />
            <Radio name="period" checked={s.periodMode === 'month'} onChange={() => s.setPeriodMode('month')}
              label="Tháng/Năm" disabled={disabled} />
            {s.periodMode === 'month' && (
              <span className="flex items-center gap-1">
                <TextInput className="w-[46px] text-center" value={s.month} disabled={disabled}
                  onChange={(e) => s.setMonth(e.target.value)} />
                <span className="text-muted">/</span>
                <TextInput className="w-[62px] text-center" value={s.year} disabled={disabled}
                  onChange={(e) => s.setYear(e.target.value)} />
              </span>
            )}
            <Radio name="period" checked={s.periodMode === 'range'} onChange={() => s.setPeriodMode('range')}
              label="Khoảng ngày" disabled={disabled} />
            {s.periodMode === 'range' && (
              <span className="flex items-center gap-1">
                <TextInput className="w-[92px] text-center" value={s.fromDate} disabled={disabled}
                  onChange={(e) => s.setFromDate(e.target.value)} />
                <span className="text-muted">→</span>
                <TextInput className="w-[92px] text-center" value={s.toDate} disabled={disabled}
                  onChange={(e) => s.setToDate(e.target.value)} />
                <span className="text-xs2 text-muted">dd/mm/yyyy, bao gồm cả hai đầu</span>
              </span>
            )}
            <span className="inline-flex items-center gap-1 text-xs2 text-muted">
              <RefreshCw className="h-3 w-3" />
              Lọc trước khi mở PPTX; theo ngày phát sinh từ Management Number
            </span>
          </span>
        </div>
      </div>
    </Card>
  )
}
