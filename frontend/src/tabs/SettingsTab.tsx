import { useEffect, useRef, useState } from 'react'
import {
  Activity,
  AlertTriangle,
  CheckCircle2,
  Cpu,
  Download,
  FileSpreadsheet,
  FolderOpen,
  Gauge,
  HardDrive,
  Loader2,
  Network,
  RefreshCw,
  Save,
  ScrollText,
  Server,
  Settings2,
  Square,
  Trash2,
  Wifi,
  WifiOff,
  Wrench,
  XCircle,
} from 'lucide-react'
import { Button, Card, Checkbox, Field, ProgressBar, TextInput, ToneChip } from '../components/ui'
import { useStore } from '../state/store'
import { LOG_LEVEL_STYLE } from '../lib/logStyle'
import { cn } from '../lib/utils'
import type { DiagnosticRow } from '../types'

type SectionKey = 'ollama' | 'options' | 'update' | 'diagnostics' | 'log'

const MENU: { key: SectionKey; label: string; icon: typeof Server; hint: string }[] = [
  { key: 'ollama', label: 'Kết nối Ollama', icon: Server, hint: 'Server, port, model, tìm trong LAN' },
  { key: 'options', label: 'Tùy chọn xử lý', icon: Settings2, hint: 'Tùy chọn xử lý lại trường tự động' },
  { key: 'update', label: 'Cập nhật', icon: Download, hint: 'Thư mục cập nhật local hoặc LAN' },
  { key: 'diagnostics', label: 'Chẩn đoán', icon: Wrench, hint: 'Ollama, template, thư mục đầu ra, QPN' },
  { key: 'log', label: 'Nhật ký', icon: ScrollText, hint: 'Nhật ký đầy đủ và file log' },
]

export function SettingsTab() {
  const [section, setSection] = useState<SectionKey>('ollama')

  return (
    <div className="flex min-h-full flex-col gap-3 p-3 lg:flex-row">
      <aside className="w-full shrink-0 lg:w-[244px]">
        <Card title="Cài đặt" dense bodyClassName="p-1.5">
          <nav className="flex flex-col gap-0.5">
            {MENU.map(({ key, label, hint, icon: Icon }) => (
              <button
                key={key}
                type="button"
                onClick={() => setSection(key)}
                className={cn(
                  'flex items-start gap-2 rounded-sm2 px-2 py-1.5 text-left',
                  section === key ? 'bg-brand-50 text-brand-700' : 'text-body hover:bg-[#f4f7fb]',
                )}
              >
                <Icon className={cn('mt-[1px] h-4 w-4 shrink-0', section === key ? 'text-brand-600' : 'text-muted')} />
                <span className="min-w-0">
                  <span className={cn('block text-base2', section === key && 'font-semibold')}>{label}</span>
                  <span className="block text-xxs text-muted">{hint}</span>
                </span>
              </button>
            ))}
          </nav>
        </Card>
      </aside>

      <div className="flex min-w-0 flex-1 flex-col gap-3">
        {section === 'ollama' && <OllamaSection />}
        {section === 'options' && <OptionsSection />}
        {section === 'update' && <UpdateSection />}
        {section === 'diagnostics' && <DiagnosticsSection />}
        {section === 'log' && <LogSection />}
      </div>
    </div>
  )
}

/* ------------------------------------------------------------------ Ollama */

function OllamaSection() {
  const s = useStore()

  const tone = s.ollama.checked === 'ok' ? 'ok' : s.ollama.checked === 'fail' ? 'warn' : 'muted'
  const icon = s.ollama.checked === 'ok' ? <Wifi className="h-3.5 w-3.5" /> : s.ollama.checked === 'fail' ? <WifiOff className="h-3.5 w-3.5" /> : null

  return (
    <>
      <Card title="Kết nối Ollama">
        <div className="max-w-[760px]">
          <Field label="Server" hint="Nhận host, host:port hoặc http://host:port. Ưu tiên kết nối cục bộ trước.">
            <div className="flex items-center gap-2">
              <TextInput
                className="max-w-[280px]"
                value={s.ollama.server}
                disabled={s.locked}
                onChange={(e) => s.setOllama({ ...s.ollama, server: e.target.value, checked: 'unchecked' })}
              />
              <span className="text-xs2 text-muted">Ollama phải chạy với OLLAMA_HOST=0.0.0.0 khi dùng trong LAN.</span>
            </div>
          </Field>
          <Field label="Port">
            <TextInput
              className="w-[84px]"
              value={s.ollama.port}
              disabled={s.locked}
              onChange={(e) => s.setOllama({ ...s.ollama, port: e.target.value, checked: 'unchecked' })}
            />
          </Field>
          <Field label="Model" hint="Chỉ dùng model để xác định vị trí nội dung; chương trình sao chép văn bản và ảnh gốc.">
            <div className="flex items-center gap-2">
              <TextInput
                className="max-w-[320px]"
                list="ollama-installed-models"
                disabled={s.locked}
                value={s.ollama.model}
                onChange={(e) => s.setOllama({ ...s.ollama, model: e.target.value, checked: 'unchecked' })}
                placeholder="Nhập tên model hoặc chọn model đã cài"
              />
              <datalist id="ollama-installed-models">
                {s.ollama.models.map((model) => <option key={model} value={model} />)}
              </datalist>
              <span className="text-xs2 text-muted">{s.ollama.models.length} model được lấy từ server Ollama</span>
            </div>
          </Field>

          <div className="mt-2 flex flex-wrap items-center gap-2">
            <Button disabled={s.locked} icon={<Activity className="h-3.5 w-3.5" />} onClick={s.checkConnection}>
              Kiểm tra kết nối
            </Button>
            <Button disabled={s.locked} icon={<RefreshCw className="h-3.5 w-3.5" />} onClick={s.refreshModels}>
              Làm mới model
            </Button>
            <Button variant="primary" disabled={s.locked} icon={<Save className="h-3.5 w-3.5" />} onClick={s.saveOllamaConfig}>
              Lưu cấu hình
            </Button>
          </div>

          <div className="mt-2.5 flex flex-col items-start gap-1.5 border-t border-lineSoft pt-2">
            <ToneChip tone={tone} icon={icon}>
              {s.ollama.checked === 'ok'
                ? 'Đã kết nối'
                : s.ollama.checked === 'fail'
                  ? 'Không kết nối được'
                  : 'Chưa kiểm tra kết nối'}
            </ToneChip>
            <p className="text-base2 text-body">{s.ollama.message}</p>
            <p className="flex gap-1.5 text-xs2 leading-[16px] text-muted">
              <AlertTriangle className="mt-[1px] h-3.5 w-3.5 shrink-0 text-warn-500" />
              {s.aiStatusText}
            </p>
          </div>
        </div>
      </Card>

      <Card title="Tìm Ollama trong mạng LAN">
        <div className="flex flex-wrap items-center gap-2">
          <Button
            icon={s.discovering ? <Loader2 className="h-3.5 w-3.5 animate-spin" /> : <Network className="h-3.5 w-3.5" />}
            disabled={s.discovering || s.locked}
            onClick={s.startDiscovery}
          >
            Tìm Ollama trong mạng LAN
          </Button>
          <Button icon={<Square className="h-3.5 w-3.5" />} disabled={!s.discovering} onClick={s.stopDiscovery}>
            Dừng tìm
          </Button>
          <span className="text-xs2 text-muted">{s.discoveryMessage}</span>
        </div>

        {s.discoveryResults.length > 0 && (
          <div className="mt-2.5 overflow-hidden rounded-sm2 border border-line">
            <table className="w-full table-fixed border-collapse text-base2">
              <thead>
                <tr className="tbl-head">
                  <th className="px-2 py-1.5 font-semibold">Server</th>
                  <th className="px-2 py-1.5 font-semibold">Port</th>
                  <th className="px-2 py-1.5 font-semibold">Model</th>
                  <th className="px-2 py-1.5 font-semibold">Phiên bản</th>
                  <th className="px-2 py-1.5 font-semibold">Ghi chú</th>
                  <th className="px-2 py-1.5" />
                </tr>
              </thead>
              <tbody>
                {s.discoveryResults.map((d) => {
                  const applied = s.appliedServer === `${d.host}:${d.port}`
                  return (
                    <tr key={d.host} className="border-b border-lineSoft last:border-b-0">
                      <td className="px-2 py-1.5 font-mono">{d.host}</td>
                      <td className="px-2 py-1.5">{d.port}</td>
                      <td className="px-2 py-1.5">{d.models}</td>
                      <td className="px-2 py-1.5">{d.version}</td>
                      <td className="px-2 py-1.5 text-muted">{d.note}</td>
                      <td className="px-2 py-1.5 text-right">
                        <Button
                          variant={applied ? 'ghost' : 'outline'}
                          disabled={applied || s.locked}
                          onClick={() => s.useServer(d.host, d.port)}
                        >
                          {applied ? 'Đang dùng server này' : 'Sử dụng server này'}
                        </Button>
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>
        )}
        <p className="mt-1.5 text-xs2 text-muted">
          Chỉ quét cổng 11434 trong dải mạng nội bộ khi người dùng bấm nút; chương trình không tự đổi server.
        </p>
      </Card>

    </>
  )
}

/* ------------------------------------------------------------------ Tùy chọn xử lý */

function OptionsSection() {
  const s = useStore()
  return (
    <Card title="Tùy chọn xử lý">
      <div className="max-w-[760px] flex flex-col gap-2">
        <Checkbox
          checked={s.force}
          disabled={s.locked}
          onChange={s.setForce}
          label="Xử lý lại báo cáo đã xử lý — ghi đè các trường tự động"
        />
        <p className="text-xs2 leading-[16px] text-muted">
          Mặc định tắt. Khi bật, các báo cáo đã xử lý sẽ được xử lý lại và ghi đè các trường do chương trình tự động
          điền. Tùy chọn này đồng bộ với tab “Danh sách báo cáo”.
        </p>
        <div className="rounded-sm2 border border-lineSoft bg-[#fafcfe] px-2.5 py-2 text-xs2 text-muted">
          Vendor name và Ngày phát sinh vẫn để trống cho người dùng điền tay; chương trình không tự điền kể cả khi bật
          tùy chọn này.
        </div>
      </div>
    </Card>
  )
}

/* ------------------------------------------------------------------ Cập nhật */

function UpdateSection() {
  const s = useStore()
  const u = s.update
  const installing = u.status === 'installing'

  return (
    <>
      <Card title="Cập nhật">
        <div className="max-w-[820px]">
          <Field label="Phiên bản hiện tại">
            <div className="flex items-center gap-2">
              <span className="font-semibold text-header">
                {s.appInfo.version || '—'} — Build {s.appInfo.build || '—'}
              </span>
              <ToneChip tone="info">development</ToneChip>
            </div>
          </Field>
          <Field label="Đường dẫn cập nhật" hint="Thư mục local hoặc LAN chứa version.json và gói ZIP. Dùng quyền truy cập Windows hiện có; chương trình không lưu mật khẩu.">
            <div className="flex items-center gap-2">
              <TextInput
                className="flex-1"
                value={u.path}
                disabled={installing || s.locked}
                onChange={(e) => s.setUpdate({ ...u, path: e.target.value })}
              />
              <Button disabled={installing || s.locked} onClick={s.chooseUpdateFolder}>
                Chọn…
              </Button>
            </div>
          </Field>
          <Field label=" ">
            <Checkbox
              checked={u.autoCheck}
              disabled={installing || s.locked}
              onChange={(v) => s.setUpdate({ ...u, autoCheck: v })}
              label="Tự động kiểm tra cập nhật khi khởi động"
            />
          </Field>

          <div className="mt-2 flex flex-wrap items-center gap-2">
            <Button
              icon={u.status === 'checking' ? <Loader2 className="h-3.5 w-3.5 animate-spin" /> : <RefreshCw className="h-3.5 w-3.5" />}
              disabled={u.status === 'checking' || installing || s.locked}
              onClick={() => s.checkUpdate(false)}
            >
              Kiểm tra cập nhật
            </Button>
            <Button
              variant="primary"
              icon={<Download className="h-3.5 w-3.5" />}
              disabled={true}
              onClick={s.installUpdate}
            >
              Cập nhật ngay
            </Button>
            <span className="text-xs2 text-muted">
              {u.autoCheck ? 'Có kiểm tra tự động khi khởi động (nhắc một lần cho mỗi bản).' : 'Đã tắt kiểm tra tự động.'}
              {' '}Việc áp dụng/cài gói cập nhật đang bị khóa trong bản tích hợp phát triển.
            </span>
          </div>

          <div className="mt-2.5 flex flex-col items-start gap-1.5 border-t border-lineSoft pt-2">
            {u.status === 'available' && <ToneChip tone="info" icon={<Download className="h-3.5 w-3.5" />}>Có bản cập nhật</ToneChip>}
            {u.status === 'installed' && <ToneChip tone="ok" icon={<CheckCircle2 className="h-3.5 w-3.5" />}>Gói đã sẵn sàng</ToneChip>}
            {u.status === 'idle' && <ToneChip tone="muted">Chưa kiểm tra</ToneChip>}
            {u.status === 'latest' && <ToneChip tone="ok" icon={<CheckCircle2 className="h-3.5 w-3.5" />}>Đang dùng bản mới nhất</ToneChip>}
            {u.status === 'error' && <ToneChip tone="error" icon={<AlertTriangle className="h-3.5 w-3.5" />}>Không thể kiểm tra cập nhật</ToneChip>}
            <p className="text-base2 text-body">{u.message}</p>
            {installing && (
              <div className="mt-1">
                <div className="flex items-center gap-2 text-xs2">
                  <span className="text-body">Trạng thái tiến trình: {u.progressStage}</span>
                  <span className="ml-auto font-semibold text-brand-600">{u.progress}%</span>
                </div>
                <ProgressBar className="mt-1" value={u.progress} />
                <p className="mt-1 text-xs2 text-muted">{u.message}</p>
              </div>
            )}
          </div>
        </div>
      </Card>

    </>
  )
}

/* ------------------------------------------------------------------ Chẩn đoán */

const DIAG_TONE: Record<DiagnosticRow['state'], { tone: 'ok' | 'warn' | 'error' | 'muted'; icon: React.ReactNode }> = {
  ok: { tone: 'ok', icon: <CheckCircle2 className="h-4 w-4 text-ok-500" /> },
  warn: { tone: 'warn', icon: <AlertTriangle className="h-4 w-4 text-warn-500" /> },
  error: { tone: 'error', icon: <XCircle className="h-4 w-4 text-danger-500" /> },
  idle: { tone: 'muted', icon: <Gauge className="h-4 w-4 text-muted" /> },
}

function DiagnosticsSection() {
  const s = useStore()
  return (
    <Card
      title="Chẩn đoán"
      actions={
        <Button
          variant="primary"
          icon={s.diagRunning ? <Loader2 className="h-3.5 w-3.5 animate-spin" /> : <Wrench className="h-3.5 w-3.5" />}
          disabled={s.diagRunning || s.locked}
          onClick={s.runDiagnostics}
        >
          {s.diagRunning ? 'Đang chạy chẩn đoán…' : 'Chạy chẩn đoán'}
        </Button>
      }
    >
      <div className="max-w-[900px]">
        <ul className="divide-y divide-lineSoft rounded-sm2 border border-line">
          {(s.diagRows.length ? s.diagRows : idleRows()).map((row) => (
            <li key={row.label} className="flex items-start gap-2.5 bg-white px-2.5 py-2">
              {DIAG_TONE[row.state].icon}
              <span className="w-[150px] shrink-0 text-base2 font-medium text-header">{row.label}</span>
              <span className={cn('flex-1 text-base2', row.state === 'idle' ? 'text-muted' : 'text-body')}>
                {row.value}
              </span>
              {row.state !== 'idle' && <ToneChip tone={DIAG_TONE[row.state].tone}>{labelOfState(row.state)}</ToneChip>}
            </li>
          ))}
        </ul>
        <p className="mt-2 text-xs2 text-muted">
          Chẩn đoán chỉ đọc: không mở PPTX, không ghi vào file kiểm chứng và không thay đổi cấu hình.
        </p>
      </div>
    </Card>
  )
}

function idleRows(): DiagnosticRow[] {
  return [
    { label: 'Ollama', value: 'Chưa chạy kiểm tra', state: 'idle' },
    { label: 'Template Excel', value: 'Chưa chạy kiểm tra', state: 'idle' },
    { label: 'Thư mục đầu ra', value: 'Chưa chạy kiểm tra', state: 'idle' },
    { label: 'Bộ dựng ảnh QPN', value: 'Chưa chạy kiểm tra', state: 'idle' },
  ]
}

function labelOfState(state: DiagnosticRow['state']) {
  return state === 'ok' ? 'Đạt' : state === 'warn' ? 'Cảnh báo' : state === 'error' ? 'Lỗi' : 'Chưa kiểm tra'
}

/* ------------------------------------------------------------------ Nhật ký */

function LogSection() {
  const s = useStore()
  const boxRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const box = boxRef.current
    if (box) box.scrollTop = box.scrollHeight
  }, [s.log])

  return (
    <Card
      title="Nhật ký"
      actions={
        <div className="flex items-center gap-2">
          <Button icon={<ScrollText className="h-3.5 w-3.5" />} onClick={s.openLogFile}>
            Mở file log
          </Button>
          <Button icon={<FolderOpen className="h-3.5 w-3.5" />} onClick={s.openLogFolder}>
            Mở thư mục log
          </Button>
          <Button icon={<Trash2 className="h-3.5 w-3.5" />} onClick={s.clearLogView}>
            Xóa phần hiển thị
          </Button>
        </div>
      }
    >
      <div ref={boxRef} className="h-[360px] overflow-auto rounded-sm2 border border-line bg-[#fbfcfe] px-3 py-2">
        {s.log.length === 0 && <p className="text-base2 text-muted">Phần hiển thị đang trống.</p>}
        {s.log.map((entry, i) => (
          <p key={i} className="whitespace-pre-wrap font-mono text-sm2 leading-[19px] text-body">
            <span className="text-muted">[{entry.time}]</span>{' '}
            <span className={cn('font-semibold', LOG_LEVEL_STYLE[entry.level])}>[{entry.level}]</span> {entry.text}
          </p>
        ))}
      </div>
      <p className="mt-2 flex flex-wrap items-center gap-x-3 gap-y-1 text-xs2 text-muted">
        <span className="inline-flex items-center gap-1">
          <HardDrive className="h-3.5 w-3.5" /> Đường dẫn log do Python backend quản lý; dùng nút “Mở file log” hoặc “Mở thư mục log”.
        </span>
        <span className="inline-flex items-center gap-1">
          <FileSpreadsheet className="h-3.5 w-3.5" /> “Xóa phần hiển thị” không xóa file log trên đĩa.
        </span>
        <span className="inline-flex items-center gap-1">
          <Cpu className="h-3.5 w-3.5" /> Log ghi bằng UTF-8, có thể mở trực tiếp khi cần gửi báo cáo lỗi.
        </span>
      </p>
    </Card>
  )
}
