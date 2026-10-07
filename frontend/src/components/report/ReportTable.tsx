import { RotateCw, Search, Trash2, Undo2 } from 'lucide-react'
import { Button, EmptyState, Select, StatusChip, TextInput } from '../ui'
import { FILTERS, useStore } from '../../state/store'
import { cn } from '../../lib/utils'
import { STAGE_LABELS } from '../../types'

export function ReportTable() {
  const s = useStore()
  const rows = s.filteredReports
  const locked = s.locked
  const allIds = rows.map((r) => r.id)
  const allChecked = allIds.length > 0 && allIds.every((id) => s.selectedIds.includes(id))

  return (
    <section className="card flex min-h-0 flex-col">
      <header className="card-head">
        <h2 className="card-title">Danh sách báo cáo</h2>
        <span className="text-xs2 text-muted">
          {rows.length}/{s.reports.length} file · hàng đợi {s.queue.length}
        </span>
      </header>

      <div className="flex flex-col gap-2 p-2.5">
        <div className="flex flex-wrap items-center gap-2">
          <span className="relative min-w-[160px] flex-1">
            <Search className="pointer-events-none absolute left-2 top-1/2 h-3.5 w-3.5 -translate-y-1/2 text-muted" />
            <TextInput
              className="pl-6"
              value={s.search}
              onChange={(e) => s.setSearch(e.target.value)}
              placeholder="Tìm theo Management Number hoặc tên file"
            />
          </span>
          <span className="flex shrink-0 items-center gap-1.5">
            <span className="whitespace-nowrap text-base2 text-body">Lọc trạng thái</span>
            <Select value={s.filter} onChange={(e) => s.setFilter(e.target.value as (typeof FILTERS)[number])}>
              {FILTERS.map((f) => (
                <option key={f} value={f}>
                  {f}
                </option>
              ))}
            </Select>
          </span>
          <Button disabled={locked} icon={<RotateCw className="h-3.5 w-3.5" />} onClick={s.scan}>
            Quét lại
          </Button>
          <Button disabled={locked} icon={<Undo2 className="h-3.5 w-3.5" />} onClick={s.restoreSelected}>
            Khôi phục
          </Button>
          <Button disabled={locked} icon={<Trash2 className="h-3.5 w-3.5" />} onClick={s.excludeSelected}>
            Xóa khỏi danh sách
          </Button>
        </div>

        <div className="min-h-[300px] flex-1 overflow-auto rounded-sm2 border border-line">
          {rows.length === 0 ? (
            s.reports.length === 0 ? (
              <EmptyState
                title="Chưa có dữ liệu"
                hint="Chưa chọn thư mục báo cáo PPTX. Chọn thư mục rồi bấm “Quét” để nạp danh sách báo cáo."
                action={
                  <Button variant="primary" onClick={s.chooseReportFolder}>
                    Chọn thư mục báo cáo
                  </Button>
                }
              />
            ) : (
              <EmptyState
                title="Không có báo cáo khớp bộ lọc"
                hint="Thử đổi từ khóa tìm kiếm hoặc trạng thái đang lọc. Danh sách đã loại thủ công vẫn xem được ở mục “File đã loại thủ công”."
                action={
                  <Button
                    onClick={() => {
                      s.setSearch('')
                      s.setFilter('Tất cả file đã quét')
                    }}
                  >
                    Xóa bộ lọc
                  </Button>
                }
              />
            )
          ) : (
            <table className="w-full table-fixed border-collapse text-base2">
              <colgroup>
                <col style={{ width: 34 }} />
                <col style={{ width: 44 }} />
                <col style={{ width: 145 }} />
                <col style={{ width: 92 }} />
                <col style={{ width: 116 }} />
                <col />
                <col style={{ width: 150 }} />
                <col style={{ width: '26%' }} />
              </colgroup>
              <thead className="sticky top-0 z-10">
                <tr className="tbl-head">
                  <th className="px-2 py-1.5">
                    <button
                      type="button"
                      onClick={() => s.toggleAllSelected(allIds)}
                      className={cn(
                        'flex h-[15px] w-[15px] items-center justify-center rounded-[2px] border',
                        allChecked ? 'border-brand-600 bg-brand-500 text-white' : 'border-[#9fb2c7] bg-white',
                      )}
                      aria-label="Chọn tất cả"
                    >
                      {allChecked && (
                        <svg viewBox="0 0 12 12" className="h-[11px] w-[11px]" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M2.5 6.2 4.8 8.5 9.5 3.6" strokeLinecap="round" strokeLinejoin="round" />
                        </svg>
                      )}
                    </button>
                  </th>
                  <th className="px-2 py-1.5 text-center font-semibold">STT</th>
                  <th className="px-2 py-1.5 font-semibold">Management Number</th>
                  <th className="px-2 py-1.5 font-semibold">Ngày phát sinh</th>
                  <th className="px-2 py-1.5 font-semibold">Vendor</th>
                  <th className="px-2 py-1.5 font-semibold">Tên file</th>
                  <th className="px-2 py-1.5 font-semibold">Trạng thái</th>
                  <th className="px-2 py-1.5 font-semibold">Cảnh báo / kết quả</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((r, index) => {
                  const active = r.id === s.activeId
                  const checked = s.selectedIds.includes(r.id)
                  const isProcessing = r.status === 'processing'
                  const attention = r.status === 'needs_review' || r.status === 'error'
                  return (
                    <tr
                      key={r.id}
                      onClick={() => s.setActiveId(r.id)}
                      className={cn(
                        'cursor-pointer border-b border-lineSoft align-top',
                        index % 2 === 1 && 'bg-[#fbfcfe]',
                        isProcessing && 'bg-brand-50',
                        r.status === 'needs_review' && 'bg-warn-50/70',
                        r.status === 'error' && 'bg-danger-50/70',
                        r.status === 'excluded' && 'text-muted',
                        active ? 'outline outline-1 -outline-offset-1 outline-brand-400' : 'hover:bg-[#f4f8fd]',
                      )}
                    >
                      <td className="px-2 py-1.5">
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation()
                            s.toggleSelected(r.id)
                          }}
                          className={cn(
                            'flex h-[15px] w-[15px] items-center justify-center rounded-[2px] border',
                            checked ? 'border-brand-600 bg-brand-500 text-white' : 'border-[#9fb2c7] bg-white',
                          )}
                          aria-label={`Chọn ${r.fileName}`}
                        >
                          {checked && (
                            <svg viewBox="0 0 12 12" className="h-[11px] w-[11px]" fill="none" stroke="currentColor" strokeWidth="2">
                              <path d="M2.5 6.2 4.8 8.5 9.5 3.6" strokeLinecap="round" strokeLinejoin="round" />
                            </svg>
                          )}
                        </button>
                      </td>
                      <td className="px-2 py-1.5 text-center text-muted">{r.stt}</td>
                      <td className={cn('truncate px-2 py-1.5', attention && 'font-semibold')}>{r.managementNumber}</td>
                      <td className="truncate px-2 py-1.5" title={r.occurrenceDate}>{r.occurrenceDate || '—'}</td>
                      <td className="truncate px-2 py-1.5" title={r.vendor}>{r.vendor || '—'}</td>
                      <td className={cn('truncate px-2 py-1.5', attention && 'font-semibold')} title={r.fileName}>
                        {r.fileName}
                      </td>
                      <td className="px-2 py-1.5">
                        <StatusChip status={r.status} />
                      </td>
                      <td className="px-2 py-1.5">
                        <div className="flex h-[34px] flex-col justify-start overflow-hidden">
                          <span className="line-clamp-1">
                            {isProcessing ? STAGE_LABELS[s.run.stage] + ' …' : r.shortResult}
                          </span>
                          {r.warning && !isProcessing && (
                            <span className="mt-0.5 line-clamp-1 text-xs2 text-warn-600">{r.warning}</span>
                          )}
                        </div>
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          )}
        </div>

        <p className="text-xs2 text-muted">
          “Xóa khỏi danh sách” chỉ loại báo cáo khỏi hàng đợi, <span className="font-medium">không xóa file nguồn</span>.
          Chọn một dòng để xem chi tiết bên phải.
        </p>
      </div>
    </section>
  )
}
