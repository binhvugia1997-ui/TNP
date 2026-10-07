export function cn(...parts: (string | false | null | undefined)[]): string {
  return parts.filter(Boolean).join(' ')
}

/** MM:SS dưới một giờ, HH:MM:SS từ một giờ trở lên (format_elapsed). */
export function formatElapsed(seconds: number): string {
  const s = Math.max(0, Math.floor(seconds))
  const h = Math.floor(s / 3600)
  const m = Math.floor((s % 3600) / 60)
  const sec = s % 60
  const two = (n: number) => String(n).padStart(2, '0')
  return h > 0 ? `${two(h)}:${two(m)}:${two(sec)}` : `${two(m)}:${two(sec)}`
}

export function formatClock(d: Date): string {
  const two = (n: number) => String(n).padStart(2, '0')
  return `${two(d.getHours())}:${two(d.getMinutes())}:${two(d.getSeconds())}`
}

export function formatStamp(d: Date): string {
  const two = (n: number) => String(n).padStart(2, '0')
  return `${two(d.getDate())}/${two(d.getMonth() + 1)}/${d.getFullYear()} ${two(d.getHours())}:${two(d.getMinutes())}`
}

export function pct(value: number, digits = 0): string {
  return `${value.toFixed(digits)}%`
}

/** So khớp tìm kiếm: Management Number / tên file / đường dẫn (row_matches_search). */
export function matchesSearch(hay: string, needle: string): boolean {
  const n = needle.trim().toLowerCase()
  if (!n) return true
  return hay.toLowerCase().includes(n)
}
