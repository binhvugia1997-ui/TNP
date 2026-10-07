import { FolderOpen, FileSpreadsheet, Wifi, WifiOff } from 'lucide-react'
import { Button } from '../ui'
import { useStore } from '../../state/store'
import { cn } from '../../lib/utils'

export function FooterBar() {
  const s = useStore()
  const { stats, run } = s
  const tone =
    s.ollama.checked === 'ok' ? 'text-ok-600' : s.ollama.checked === 'fail' ? 'text-warn-600' : 'text-muted'

  return (
    <div className="card flex flex-wrap items-center gap-x-4 gap-y-2 px-3 py-2">
      <Button
        icon={<FileSpreadsheet className="h-3.5 w-3.5" />}
        disabled={s.locked}
        onClick={s.openOutputFile}
      >
        Mở file kết quả
      </Button>
      <Button icon={<FolderOpen className="h-3.5 w-3.5" />} disabled={s.locked} onClick={s.openOutputFolder}>
        Mở thư mục kết quả
      </Button>

      <span className="text-xs2 text-body">
        Tổng: <span className="font-semibold">{stats.total}</span> · Hoàn thành:{' '}
        <span className="font-semibold">{stats.completed}</span> · Đang xử lý:{' '}
        <span className="font-semibold">{run.status === 'processing' ? 1 : 0}</span> · Còn lại:{' '}
        <span className="font-semibold">{s.queue.length}</span>
      </span>

      <span className="ml-auto flex items-center gap-3">
        <span className="text-xs2 text-muted">
          Kết quả:{' '}
          <span className="font-mono text-body">{s.output}</span>
        </span>
        <span className={cn('inline-flex items-center gap-1.5 text-xs2 font-medium', tone)}>
          {s.ollama.checked === 'fail' ? <WifiOff className="h-3.5 w-3.5" /> : <Wifi className="h-3.5 w-3.5" />}
          <span className="h-2 w-2 rounded-full bg-current" />
          {s.ollamaIndicator.text}
        </span>
      </span>
    </div>
  )
}
