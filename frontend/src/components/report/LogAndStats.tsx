import { useEffect, useRef } from 'react'
import { AlertTriangle, CheckCircle2, ChevronDown, MinusCircle, XCircle } from 'lucide-react'
import { Card, StatCard } from '../ui'
import { useStore } from '../../state/store'
import { LOG_LEVEL_STYLE } from '../../lib/logStyle'
import { cn } from '../../lib/utils'

export function LogCard() {
  const s = useStore()
  const boxRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const box = boxRef.current
    if (box && !s.logCollapsed) box.scrollTop = box.scrollHeight
  }, [s.log, s.logCollapsed])

  const last = s.log[s.log.length - 1]

  return (
    <Card
      title="Nhật ký xử lý"
      actions={
        <div className="flex items-center gap-2">
          <span className="text-xs2 text-muted">Nhật ký đầy đủ nằm trong tab Cài đặt → Nhật ký</span>
          <button
            type="button"
            onClick={() => s.setLogCollapsed(!s.logCollapsed)}
            className="inline-flex items-center gap-1 text-xs2 text-brand-700 hover:underline"
          >
            <ChevronDown className={cn('h-3.5 w-3.5 transition-transform', s.logCollapsed && '-rotate-90')} />
            {s.logCollapsed ? 'Mở rộng' : 'Thu gọn'}
          </button>
        </div>
      }
    >
      {s.logCollapsed ? (
        <p className="truncate px-1 font-mono text-xs2 text-muted">
          {last ? `[${last.time}] [${last.level}] ${last.text}` : 'Nhật ký trống.'}
        </p>
      ) : (
        <div ref={boxRef} className="h-[136px] overflow-auto rounded-sm2 border border-lineSoft bg-[#fbfcfe] px-2 py-1.5">
          {s.log.length === 0 && <p className="text-xs2 text-muted">Chưa có dòng nhật ký nào.</p>}
          {s.log.map((entry, i) => (
            <p key={i} className="whitespace-pre-wrap font-mono text-xs2 leading-[17px] text-body">
              <span className="text-muted">[{entry.time}]</span>{' '}
              <span className={cn('font-semibold', LOG_LEVEL_STYLE[entry.level])}>[{entry.level}]</span> {entry.text}
            </p>
          ))}
        </div>
      )}
    </Card>
  )
}

export function StatsCard() {
  const s = useStore()
  const { stats } = s
  return (
    <Card
      title="Thống kê"
      actions={<span className="text-xs2 text-muted">Tổng số báo cáo: {stats.total}</span>}
    >
      <div className="grid grid-cols-2 gap-2">
        <StatCard tone="ok" icon={<CheckCircle2 className="h-4 w-4" />} value={stats.completed} label="Hoàn thành" />
        <StatCard
          tone="warn"
          icon={<AlertTriangle className="h-4 w-4" />}
          value={stats.needs_review}
          label="Cần kiểm tra"
        />
        <StatCard tone="error" icon={<XCircle className="h-4 w-4" />} value={stats.error} label="Lỗi" />
        <StatCard tone="muted" icon={<MinusCircle className="h-4 w-4" />} value={stats.skipped} label="Bỏ qua" />
      </div>
      <p className="mt-2 text-xs2 leading-[16px] text-muted">
        Số liệu lấy trực tiếp từ danh sách báo cáo: {stats.completed} hoàn thành · {stats.needs_review} cần kiểm tra ·{' '}
        {stats.error} lỗi · {stats.skipped} bỏ qua · {s.queue.length} đang chờ xử lý.
      </p>
    </Card>
  )
}
