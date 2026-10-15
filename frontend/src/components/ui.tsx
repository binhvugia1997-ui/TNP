import type { ButtonHTMLAttributes, InputHTMLAttributes, ReactNode, SelectHTMLAttributes } from 'react'
import { AlertTriangle, CheckCircle2, MinusCircle, X, XCircle } from 'lucide-react'
import { cn } from '../lib/utils'
import { STATUS_LABEL, type StatusKey } from '../types'

/* ------------------------------------------------------------------ Card */

export function Card({
  title,
  icon,
  actions,
  children,
  className,
  bodyClassName,
  dense,
}: {
  title?: ReactNode
  icon?: ReactNode
  actions?: ReactNode
  children: ReactNode
  className?: string
  bodyClassName?: string
  dense?: boolean
}) {
  return (
    <section className={cn('card flex min-h-0 flex-col', className)}>
      {title !== undefined && (
        <header className="card-head">
          {icon && <span className="text-brand-500">{icon}</span>}
          <h2 className="card-title mr-auto">{title}</h2>
          {actions}
        </header>
      )}
      <div className={cn(dense ? 'p-0' : 'p-3', 'min-h-0 flex-1', bodyClassName)}>{children}</div>
    </section>
  )
}

/* ------------------------------------------------------------------ Button */

type ButtonProps = ButtonHTMLAttributes<HTMLButtonElement> & {
  variant?: 'outline' | 'primary' | 'danger' | 'ghost'
  icon?: ReactNode
  active?: boolean
}

export function Button({ variant = 'outline', icon, children, className, active, ...rest }: ButtonProps) {
  return (
    <button
      type="button"
      {...rest}
      className={cn(
        'btn',
        variant === 'primary' && 'btn-primary',
        variant === 'outline' && 'btn-outline',
        variant === 'danger' && 'btn-danger',
        variant === 'ghost' && 'btn-ghost',
        active && 'ring-1 ring-brand-400',
        className,
      )}
    >
      {icon}
      {children}
    </button>
  )
}

/* ------------------------------------------------------------------ Form */

export function TextInput({ className, ...rest }: InputHTMLAttributes<HTMLInputElement>) {
  return <input type="text" {...rest} className={cn('field', className)} />
}

export function Select({ className, children, ...rest }: SelectHTMLAttributes<HTMLSelectElement>) {
  return (
    <select {...rest} className={cn('field cursor-pointer pr-6', className)}>
      {children}
    </select>
  )
}

export function Checkbox({
  checked,
  onChange,
  label,
  disabled,
  className,
}: {
  checked: boolean
  onChange: (v: boolean) => void
  label: ReactNode
  disabled?: boolean
  className?: string
}) {
  return (
    <label
      className={cn(
        'inline-flex select-none items-center gap-2 text-base2',
        disabled ? 'cursor-not-allowed text-muted' : 'cursor-pointer text-body',
        className,
      )}
    >
      <span
        className={cn(
          'flex h-[15px] w-[15px] items-center justify-center rounded-[2px] border',
          checked ? 'border-brand-600 bg-brand-500 text-white' : 'border-[#9fb2c7] bg-white',
          disabled && 'opacity-60',
        )}
      >
        {checked && (
          <svg viewBox="0 0 12 12" className="h-[11px] w-[11px]" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M2.5 6.2 4.8 8.5 9.5 3.6" strokeLinecap="round" strokeLinejoin="round" />
          </svg>
        )}
      </span>
      <input
        type="checkbox"
        className="sr-only"
        checked={checked}
        disabled={disabled}
        onChange={(e) => onChange(e.target.checked)}
      />
      {label}
    </label>
  )
}

export function Radio({
  checked,
  onChange,
  label,
  name,
  disabled,
}: {
  checked: boolean
  onChange: () => void
  label: ReactNode
  name: string
  disabled?: boolean
}) {
  return (
    <label
      className={cn(
        'inline-flex select-none items-center gap-1.5 text-base2',
        disabled ? 'cursor-not-allowed text-muted' : 'cursor-pointer text-body',
      )}
    >
      <span
        className={cn(
          'flex h-[14px] w-[14px] items-center justify-center rounded-full border',
          checked ? 'border-brand-600' : 'border-[#9fb2c7]',
          disabled && 'opacity-60',
        )}
      >
        {checked && <span className="h-[7px] w-[7px] rounded-full bg-brand-500" />}
      </span>
      <input type="radio" name={name} className="sr-only" checked={checked} disabled={disabled} onChange={onChange} />
      {label}
    </label>
  )
}

export function Field({ label, children, hint }: { label: ReactNode; children: ReactNode; hint?: ReactNode }) {
  return (
    <div className="flex items-start gap-3 py-[3px]">
      <div className="w-[132px] shrink-0 pt-[3px] text-base2 text-body">{label}</div>
      <div className="min-w-0 flex-1">
        {children}
        {hint && <div className="mt-1 text-xs2 leading-[16px] text-muted">{hint}</div>}
      </div>
    </div>
  )
}

export function KeyValue({ label, children }: { label: ReactNode; children: ReactNode }) {
  return (
    <div className="flex items-start gap-2 py-[3px] text-base2">
      <div className="w-[104px] shrink-0 text-muted">{label}</div>
      <div className="min-w-0 flex-1 break-words text-body">{children}</div>
    </div>
  )
}

/* ------------------------------------------------------------------ Status chip */

const STATUS_STYLE: Record<StatusKey, { cls: string; icon: ReactNode }> = {
  completed: { cls: 'border-ok-500/30 bg-ok-50 text-ok-600', icon: <CheckCircle2 className="h-3.5 w-3.5" /> },
  processing: { cls: 'border-brand-400/40 bg-brand-50 text-brand-700', icon: <span className="h-2 w-2 animate-pulse rounded-full bg-brand-500" /> },
  needs_review: { cls: 'border-warn-500/30 bg-warn-50 text-warn-600', icon: <AlertTriangle className="h-3.5 w-3.5" /> },
  error: { cls: 'border-danger-500/30 bg-danger-50 text-danger-600', icon: <XCircle className="h-3.5 w-3.5" /> },
  // PROMPT-024R: user cancellation — neutral grey, never the red "Lỗi" treatment.
  cancelled: { cls: 'border-line bg-[#eef1f5] text-[#4a6280]', icon: <MinusCircle className="h-3.5 w-3.5" /> },
  waiting: { cls: 'border-line bg-[#f4f7fb] text-muted', icon: <MinusCircle className="h-3.5 w-3.5" /> },
  new_row: { cls: 'border-brand-200 bg-brand-50 text-brand-700', icon: <MinusCircle className="h-3.5 w-3.5" /> },
  skipped: { cls: 'border-line bg-[#f4f7fb] text-muted', icon: <MinusCircle className="h-3.5 w-3.5" /> },
  outside_period: { cls: 'border-line bg-[#f4f7fb] text-muted', icon: <MinusCircle className="h-3.5 w-3.5" /> },
  source_duplicate: { cls: 'border-line bg-[#f4f7fb] text-muted', icon: <MinusCircle className="h-3.5 w-3.5" /> },
  fast_skip: { cls: 'border-line bg-[#f4f7fb] text-muted', icon: <MinusCircle className="h-3.5 w-3.5" /> },
  excluded: { cls: 'border-line bg-[#f4f7fb] text-muted', icon: <MinusCircle className="h-3.5 w-3.5" /> },
}

export function StatusChip({ status, label, className }: { status: StatusKey; label?: string; className?: string }) {
  const style = STATUS_STYLE[status]
  return (
    <span className={cn('chip', style.cls, className)}>
      {style.icon}
      {label ?? STATUS_LABEL[status]}
    </span>
  )
}

export function ToneChip({
  tone,
  children,
  icon,
}: {
  tone: 'ok' | 'warn' | 'error' | 'info' | 'muted'
  children: ReactNode
  icon?: ReactNode
}) {
  const map = {
    ok: 'border-ok-500/30 bg-ok-50 text-ok-600',
    warn: 'border-warn-500/30 bg-warn-50 text-warn-600',
    error: 'border-danger-500/30 bg-danger-50 text-danger-600',
    info: 'border-brand-200 bg-brand-50 text-brand-700',
    muted: 'border-line bg-[#f4f7fb] text-muted',
  } as const
  return (
    <span className={cn('chip', map[tone])}>
      {icon}
      {children}
    </span>
  )
}

/* ------------------------------------------------------------------ Progress */

export function ProgressBar({ value, className }: { value: number; className?: string }) {
  return (
    <div className={cn('h-[6px] w-full overflow-hidden rounded-full border border-line bg-[#eef2f7]', className)}>
      <div
        className="h-full rounded-full bg-brand-500 transition-[width] duration-150 ease-linear"
        style={{ width: `${Math.max(0, Math.min(100, value))}%` }}
      />
    </div>
  )
}

export function StatCard({
  icon,
  value,
  label,
  tone,
}: {
  icon: ReactNode
  value: number | string
  label: string
  tone: 'info' | 'ok' | 'warn' | 'error' | 'muted'
}) {
  const ring = {
    info: 'bg-brand-50 text-brand-600 border-brand-200',
    ok: 'bg-ok-50 text-ok-500 border-ok-500/25',
    warn: 'bg-warn-50 text-warn-500 border-warn-500/25',
    error: 'bg-danger-50 text-danger-500 border-danger-500/25',
    muted: 'bg-[#f4f7fb] text-muted border-line',
  } as const
  return (
    <div className="flex items-center gap-2.5 rounded-sm2 border border-line bg-white px-2.5 py-2">
      <span className={cn('flex h-8 w-8 items-center justify-center rounded-sm2 border', ring[tone])}>{icon}</span>
      <span className="min-w-0">
        <span className="block text-[19px] font-semibold leading-[22px] text-header">{value}</span>
        <span className="block truncate text-xs2 text-muted">{label}</span>
      </span>
    </div>
  )
}

/* ------------------------------------------------------------------ Modal */

export function Modal({
  open,
  title,
  subtitle,
  onClose,
  children,
  footer,
  width = 'max-w-3xl',
}: {
  open: boolean
  title: ReactNode
  subtitle?: ReactNode
  onClose: () => void
  children: ReactNode
  footer?: ReactNode
  width?: string
}) {
  if (!open) return null
  return (
    <div className="fixed inset-0 z-40 flex items-center justify-center bg-[#0f2b4c]/35 p-6" onClick={onClose}>
      <div
        className={cn('flex max-h-[92vh] w-full flex-col overflow-hidden rounded-sm2 border border-line bg-white shadow-pop', width)}
        onClick={(e) => e.stopPropagation()}
      >
        <header className="flex items-start gap-2 border-b border-line px-3 py-2">
          <div className="mr-auto min-w-0">
            <h3 className="text-base2 font-semibold text-header">{title}</h3>
            {subtitle && <p className="truncate text-xs2 text-muted">{subtitle}</p>}
          </div>
          <button type="button" onClick={onClose} className="rounded-sm2 p-1 text-muted hover:bg-[#f2f5f9] hover:text-body">
            <X className="h-4 w-4" />
          </button>
        </header>
        <div className="min-h-0 flex-1 overflow-auto px-3 py-2.5">{children}</div>
        {footer && <footer className="flex items-center gap-2 border-t border-line bg-[#fafcfe] px-3 py-2">{footer}</footer>}
      </div>
    </div>
  )
}

/* ------------------------------------------------------------------ Misc */

export function EmptyState({ title, hint, action }: { title: string; hint: string; action?: ReactNode }) {
  return (
    <div className="flex h-full min-h-[180px] flex-col items-center justify-center gap-2 px-4 py-8 text-center">
      <div className="flex h-11 w-11 items-center justify-center rounded-full border border-line bg-[#f4f7fb] text-brand-500">
        <AlertTriangle className="h-5 w-5" />
      </div>
      <p className="text-base2 font-semibold text-header">{title}</p>
      <p className="max-w-[420px] text-xs2 text-muted">{hint}</p>
      {action}
    </div>
  )
}
