import { useState } from 'react'
import { FileSpreadsheet, FolderOpen } from 'lucide-react'
import { Button, Modal } from './ui'
import { cn } from '../lib/utils'

export type PathChoice = { path: string; note: string }

/** Hộp chọn đường dẫn mô phỏng (thao tác chọn thư mục/file chỉ là giả lập trong bản mẫu). */
export function PathPicker({
  open,
  title,
  choices,
  onClose,
  onPick,
  kind,
}: {
  open: boolean
  title: string
  choices: PathChoice[]
  onClose: () => void
  onPick: (path: string) => void
  kind: 'folder' | 'file'
}) {
  const [hover, setHover] = useState<string | null>(null)
  const Icon = kind === 'folder' ? FolderOpen : FileSpreadsheet

  return (
    <Modal
      open={open}
      title={title}
      subtitle="Mô phỏng hộp thoại chọn đường dẫn của Windows — bản mẫu không truy cập ổ đĩa thật."
      onClose={onClose}
      width="max-w-2xl"
      footer={
        <>
          <span className="mr-auto text-xs2 text-muted">Chọn một đường dẫn rồi bấm “Chọn đường dẫn này”.</span>
          <Button onClick={onClose}>Hủy</Button>
          <Button
            variant="primary"
            disabled={!hover}
            onClick={() => {
              if (hover) {
                onPick(hover)
                onClose()
              }
            }}
          >
            Chọn đường dẫn này
          </Button>
        </>
      }
    >
      <ul className="divide-y divide-lineSoft rounded-sm2 border border-line">
        {choices.map((c) => (
          <li key={c.path}>
            <button
              type="button"
              onClick={() => setHover(c.path)}
              onDoubleClick={() => {
                onPick(c.path)
                onClose()
              }}
              className={cn(
                'flex w-full items-center gap-2 px-2.5 py-2 text-left',
                hover === c.path ? 'bg-brand-50' : 'bg-white hover:bg-[#f6f9fd]',
              )}
            >
              <Icon className="h-4 w-4 shrink-0 text-brand-500" />
              <span className="min-w-0 flex-1">
                <span className="block truncate font-mono text-base2 text-body">{c.path}</span>
                <span className="block text-xs2 text-muted">{c.note}</span>
              </span>
            </button>
          </li>
        ))}
      </ul>
    </Modal>
  )
}
