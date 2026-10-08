import { FileSearch, FolderCog, GraduationCap } from 'lucide-react'
import { cn } from '../lib/utils'
import { useStore } from '../state/store'
import type { TabKey } from '../types'

const TABS: { key: TabKey; label: string; icon: typeof FileSearch }[] = [
  { key: 'reports', label: 'Danh sách báo cáo', icon: FileSearch },
  { key: 'settings', label: 'Cài đặt', icon: FolderCog },
  { key: 'learning', label: 'Học cải tiến', icon: GraduationCap },
]

export function AppHeader() {
  const { tab, setTab, title, appInfo } = useStore()
  /* PROMPT-027 §50: the running build must be identifiable AT A GLANCE during Windows acceptance, so the
     header shows version AND build together (for example "vX.Y.Z · Build NNN").  APP_TITLE already embeds
     the bare version, so the h1 shows the product NAME and the badge carries version + build — no
     duplication.  Both values come from the single authoritative Python source via the bridge; this file
     must never hardcode a version or build literal of its own. */
  const productName = appInfo.name || title
  const versionText = appInfo.version
    ? `v${appInfo.version}${appInfo.build ? ` · Build ${appInfo.build}` : ''}`
    : ''

  return (
    <header className="shrink-0 border-b border-line bg-app">
      <div className="flex items-center gap-2.5 px-3 py-2">
        <span className="flex h-[26px] w-[26px] items-center justify-center rounded-sm2 bg-brand-500 text-white">
          <svg viewBox="0 0 24 24" className="h-[15px] w-[15px]" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z" strokeLinejoin="round" />
            <path d="M14 3v5h5" strokeLinejoin="round" />
            <circle cx="11.5" cy="14.5" r="2.6" />
            <path d="m13.6 16.6 2.4 2.4" strokeLinecap="round" />
          </svg>
        </span>
        <h1 className="text-base2 font-semibold text-header">{productName}</h1>
        {versionText && (
          <span
            data-testid="app-version-badge"
            className="shrink-0 rounded-sm2 border border-line bg-white/70 px-1.5 py-[1px] text-xxs font-medium tabular-nums text-muted"
            title="Phiên bản đang chạy (nguồn duy nhất: Python backend)"
          >
            {versionText}
          </span>
        )}
        <span className="ml-1 hidden text-xs2 text-muted sm:inline">PPTX → Bảng kiểm chứng</span>
      </div>

      <nav className="flex items-end gap-1 px-2" role="tablist" aria-label="Khu vực chính">
        {TABS.map(({ key, label, icon: Icon }) => {
          const active = tab === key
          return (
            <button
              key={key}
              type="button"
              role="tab"
              aria-selected={active}
              onClick={() => setTab(key)}
              className={cn(
                'relative -mb-px inline-flex h-[32px] items-center gap-2 rounded-t-sm2 border px-4 text-base2',
                active
                  ? 'border-brand-600 bg-brand-500 font-semibold text-white'
                  : 'border-line border-b-transparent bg-white/70 text-header hover:bg-white',
              )}
            >
              <Icon className={cn('h-[15px] w-[15px]', active ? 'text-white' : 'text-brand-600')} />
              {label}
            </button>
          )
        })}
      </nav>
    </header>
  )
}
