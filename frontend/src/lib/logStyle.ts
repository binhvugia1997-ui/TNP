import type { LogEntry } from '../types'

export const LOG_LEVEL_STYLE: Record<LogEntry['level'], string> = {
  INFO: 'text-[#1d63b8]',
  WARN: 'text-[#b45309]',
  ERROR: 'text-[#b91c1c]',
  DEBUG: 'text-[#5b7290]',
}
