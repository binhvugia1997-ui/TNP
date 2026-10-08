import { AlertTriangle, CheckCircle2, X, XCircle } from 'lucide-react'
import { AppHeader } from './components/AppHeader'
import { useStore } from './state/store'
import { LearningTab } from './tabs/LearningTab'
import { ReportListTab } from './tabs/ReportListTab'
import { SettingsTab } from './tabs/SettingsTab'
import { cn } from './lib/utils'

function messageCategory(code: string) {
  if (code === 'EXCEL_LOCKED') return 'Excel đang bị khóa'
  if (code === 'BRIDGE_UNAVAILABLE') return 'Python bridge chưa sẵn sàng'
  if (code === 'BRIDGE_METHOD_MISSING' || code === 'BRIDGE_CONTRACT_ERROR') return 'Lỗi hợp đồng Python bridge'
  if (code === 'BUSY') return 'Ứng dụng đang bận'
  if (code === 'WORKER_FAILED') return 'Lỗi tiến trình xử lý'
  if (code.startsWith('OLLAMA')) return 'Ollama không khả dụng'
  if (code.includes('INVALID') || code.includes('VALIDATION')) return 'Cần kiểm tra thông tin'
  if (code === 'INTERNAL_ERROR') return 'Thao tác không thành công'
  return code ? 'Thao tác không thành công' : 'Thông báo'
}

export default function App() {
  const s = useStore()
  const failed = Boolean(s.errorCode)
  const warning = s.errorCode === 'EXCEL_LOCKED' || s.errorCode === 'BUSY' || s.errorCode === 'BRIDGE_UNAVAILABLE'

  return (
    <div className="flex h-full min-h-0 flex-col bg-app">
      <AppHeader />
      {!s.connected && !s.errorCode.startsWith('BRIDGE_') && (
        <div className="flex items-center gap-2 border-b border-warn-500/30 bg-warn-50 px-3 py-1.5 text-xs2 text-warn-600" role="status">
          <AlertTriangle className="h-3.5 w-3.5 shrink-0" />
          <span>Chưa kết nối ứng dụng Python. Chế độ demo không được bật; hãy khởi chạy bằng <code>python -m app.desktop</code>.</span>
        </div>
      )}
      {s.errorMessage && (
        <div
          className={cn(
            'flex items-start gap-2 border-b px-3 py-1.5 text-xs2',
            failed ? warning ? 'border-warn-500/30 bg-warn-50 text-warn-600' : 'border-danger-500/30 bg-danger-50 text-danger-600'
              : 'border-brand-200 bg-brand-50 text-brand-700',
          )}
          role={failed ? 'alert' : 'status'}
        >
          {failed ? warning ? <AlertTriangle className="mt-[1px] h-3.5 w-3.5 shrink-0" />
            : <XCircle className="mt-[1px] h-3.5 w-3.5 shrink-0" />
            : <CheckCircle2 className="mt-[1px] h-3.5 w-3.5 shrink-0" />}
          <span className="min-w-0 flex-1">
            <span className="font-semibold">{messageCategory(s.errorCode)}{s.errorCode ? ` · ${s.errorCode}` : ''}: </span>
            {s.errorMessage}
          </span>
          <button type="button" onClick={s.clearError} aria-label="Đóng thông báo" className="shrink-0 rounded-sm2 p-0.5 hover:bg-black/5">
            <X className="h-3.5 w-3.5" />
          </button>
        </div>
      )}
      <main className="min-h-0 flex-1 overflow-auto">
        {s.tab === 'reports' && <ReportListTab />}
        {s.tab === 'settings' && <SettingsTab />}
        {s.tab === 'learning' && <LearningTab />}
      </main>
    </div>
  )
}
